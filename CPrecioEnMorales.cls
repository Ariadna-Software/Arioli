VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrecioEnMorales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'------------------------------------------------------------------
'------------------------------------------------------------------
'
'Cuando desde la exportador generamos un pedido en la de produccion
' los precios los saca desde esta rutina
'

'Esto es para cuando AVAB crea un pedido en MORALES
'                   ********************************
'
'
'Modulo para Calcula el Precio a Aplicar a las lineas de una Oferta, Pedido o Albaran
' Y los descuentos correspondientes si se aplica descuento
'TABLAS: slista, sprees, sdtofm

'------------------------------------------------
'Variables del modulo CPreciosFact
'------------------------------------------------

Private mvarDtoPermitidMor As Boolean
Private mvarcodartiMor As String
Private mvarCodListMor As String
Private mvarCodClieMor As String
Private mvarCodFamiMor As String
Private mvarCodMarcMor As String

Private mvarDto1Mor As String
Private mvardescuento2 As String


'------------------------------------------------
'Propiedades del modulo CPreciosFact
'------------------------------------------------

Public Property Let DtoPermitidMor(ByVal vData As Boolean)
     mvarDtoPermitidMor = vData
End Property

Public Property Get DtoPermitidMor() As Boolean
     DtoPermitidMor = mvarDtoPermitidMor
End Property

Public Property Let CodigoArticMor(ByVal vData As String)
     mvarcodartiMor = vData
End Property

Public Property Get CodigoArticMor() As String
     CodigoArticMor = mvarcodartiMor
End Property

Public Property Let CodigoListaMor(ByVal vData As String)
     mvarCodListMor = vData
End Property

Public Property Get CodigoListaMor() As String
     CodigoListaMor = mvarCodListMor
End Property


Public Property Let CodigoClienMor(ByVal vData As String)
     mvarCodClieMor = vData
End Property

Public Property Get CodigoClienMor() As String
     CodigoClienMor = mvarCodClieMor
End Property

Public Property Let CodigoFamiaMor(ByVal vData As String)
     mvarCodFamiMor = vData
End Property

Public Property Get CodigoFamiaMor() As String
     CodigoFamiaMor = mvarCodFamiMor
End Property

Public Property Let CodigoMarcaMor(ByVal vData As String)
     mvarCodMarcMor = vData
End Property

Public Property Get CodigoMarcaMor() As String
     CodigoMarcaMor = mvarCodMarcMor
End Property

Public Property Let Descuento1(ByVal vData As String)
     mvarDto1Mor = vData
End Property

Public Property Get Descuento1() As String
     Descuento1 = mvarDto1Mor
End Property


Public Property Let Descuento2(ByVal vData As String)
     mvardescuento2 = vData
End Property

Public Property Get Descuento2() As String
     Descuento2 = mvardescuento2
End Property

'------------------------------------------------
'Metodos del modulo CPreciosFact
'------------------------------------------------

Public Function ObtenerPrecioMor(PCaja As Boolean, FechaOfe As String, OrigP As String) As String
'PCaja -> IN: si vale true se obiene el precio de unidad si se vende por caja sino se obtiene precio
'         OUT: si sale con precio de pvp del articulo se pone a False, independientemente de que se
'              pueda vender por cajas o no
'FechaOfe -> IN: Fecha de la Oferta para comprobar si la oferta esta en los periodos de promocion, etc...
'OrigP -> OUT: Aqui se devolvera con que valor precio se sale de la función, e.d.:
'               P=sale con precio de Promoción
'               E=sale con Precio Especial
'               etc....
'RETURN: valor del precio
Dim PrecioFact As String
Dim Dto1Fact As String
Dim Dto2Fact As String
On Error GoTo EObtenerPre
   
   
   
    'Abril 2009
    '-------------------------------------------------------------------------
    'Primero comprobaremos en las TOs
    
    PrecioFact = ObtenerTO(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el Precio
         
         OrigP = "O" 'Indica que sale con precio de Promocion
     
    Else
        'COMPROBAR PROMOCIONES
        '==============================================================
        'Buscar Precio en la tabla spromo si encuentro valor salgo con ese precio
        PrecioFact = ObtenerPromocion(FechaOfe, PCaja)
        If PrecioFact <> "" Then 'Ha obtenido el Precio
            If DtoPermitidMor Then
            'Buscar en la tabla sdtofm los descuentos
                Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
            End If
            OrigP = "P" 'Indica que sale con precio de Promocion
            
        Else 'No encuentra Precio en spromo y Buscar en Sprees(Tabla Precios Especiales)
            'COMPROBAR PRECIO ESPECIAL
            '===========================================================================
            PrecioFact = ObtenerPrecioEsp(FechaOfe, PCaja, Dto1Fact)
            Dto2Fact = "0"
            If PrecioFact <> "" Then 'Ha obtenido el precio
                If DtoPermitidMor And Dto1Fact = "" Then
                'Buscar en la tabla sdtofm los descuentos
                    Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
                End If
                OrigP = "E" 'Indica que sale con Precio Especial
            Else 'No encuentra precio en sprees y Buscar en slista(Tabla Tarifas Articulos)
                'COMPROBAR TARIFAS ARTICULOS
                '=======================================================================
                PrecioFact = ObtenerTarifaArt(FechaOfe, PCaja)
                If PrecioFact <> "" Then 'Ha obtenido el precio
                    If DtoPermitidMor Then
                        'Buscar en la tabla sdtofm los descuentos
                        Dto1Fact = ObtenerDescuentos(FechaOfe, Dto2Fact, PCaja)
                    End If
                    OrigP = "T"
                Else 'Buscar en sartic el precio del articulo (preciove)
                    PrecioFact = ObtenerPrecioVenta
                    PCaja = False
                    OrigP = "A"
                    'Nuevo###. David 14 Mayo 2008
                    'Permitimos descuento
                    DtoPermitidMor = True
                End If
            End If
        End If 'de promocion
    End If 'de TO
    ObtenerPrecioMor = ComprobarCero(PrecioFact)
    Descuento1 = ComprobarCero(Dto1Fact)
    Descuento2 = ComprobarCero(Dto2Fact)
    
EObtenerPre:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Precio", Err.Description
End Function



Private Function ObtenerPromocion(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una promocion en la tabla: spromo (Promociones Tarifas)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerPromo

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM ariges" & EmprMorales & ".spromo "
    SQL = SQL & " WHERE codartic=" & DBSet(mvarcodartiMor, "T") & " AND codlista=" & CodigoListaMor
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
    'Existe una promocion para ese articulo y tarifa, comprobar fechas
        If EntreFechas(DBLet(Rs!FechaIni), FechaOfe, DBLet(Rs!FechaFin)) Then
        'Dentro Periodo Promocion Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(Rs!precioa1))
            Else
                Precio = DBLet(Rs!precioac)
            End If
        ElseIf DBLet(Rs!fechain1) = "" And DBLet(Rs!fechafi1) = "" Then
            Precio = ""
        ElseIf EntreFechas(DBLet(Rs!fechain1), FechaOfe, DBLet(Rs!fechafi1)) Then
        'DEntro Periodo Promocion Nueva
            If PorCaja Then
                Precio = ComprobarCero(DBLet(Rs!precion1))
            Else
                Precio = ComprobarCero(DBLet(Rs!precionu))
            End If
        End If
        DtoPermitidMor = CBool(Rs!dtopermi)
    Else 'No existe Promocion salir y pasar al siguiente paso
        Precio = ""
        DtoPermitidMor = False
    End If
    Rs.Close
    Set Rs = Nothing
    ObtenerPromocion = Precio
    
EObtenerPromo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Promoción", Err.Description
End Function


Private Function ObtenerDescuentos(FechaOfe As String, Dto2Fact, PorCaja As Boolean) As String
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String
    
    vCodMarca = "Codmarca"
    CodigoFamiaMor = DevuelveDesdeBDNew(conAri, "ariges" & EmprMorales & ".sartic", "Codfamia", "codartic", CodigoArticMor, "T", vCodMarca)
    CodigoMarcaMor = vCodMarca
    If PorCaja Then 'Obtener Descuentos Por Caja
        CampoDto1 = "dtocaja1"
        campoDto2 = "dtocaja2"
    Else 'Obtener Descuentos por Unidad
        CampoDto1 = "dtoline1"
        campoDto2 = "dtoline2"
    End If
    
    'Obtener Descuento
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & ", codlista FROM ariges" & EmprMorales & ".sdtofm "
    SQL = SQL & " WHERE codclien=" & CodigoClienMor
    SQL = SQL & " AND (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND ((codfamia=" & CodigoFamiaMor & " AND codmarca=" & CodigoMarcaMor & ")"
    SQL = SQL & " OR (codfamia=" & CodigoFamiaMor & " AND codmarca=0)"
    SQL = SQL & " OR (codfamia=0 AND codmarca=" & CodigoMarcaMor & ")"
    SQL = SQL & " OR (codfamia=0 AND codmarca=0))"
    If vParamAplic.OrdenDtos = 0 Then
        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
    Else
        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentos = DBLet(Rs.Fields(0).Value)
        Dto2Fact = DBLet(Rs.Fields(1).Value)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentos = ""
        Dto2Fact = ""
    End If
    Rs.Close
    Set Rs = Nothing
End Function



Private Function ObtenerPrecioEsp(FechaOfe As String, PorCaja As Boolean, Dto As String) As String
'Comprueba si existe un Precio Especial en la tabla: sprees (Precios Especiales)
'Dim devuelve As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerPreEsp

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM ariges" & EmprMorales & ".sprees "
    SQL = SQL & " WHERE codclien=" & CodigoClienMor & " AND codartic=" & DBSet(CodigoArticMor, "T")

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
    'Existe una promocion, y salimos con el precio actual o nuevo
        DtoPermitidMor = CBool(Rs!dtopermi)
        If IsNull(Rs!fechanue) Or (EsFechaIgualPosterior(FechaOfe, DBLet(Rs!fechanue), False)) Then
        'Precio Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(Rs!precioa1))
            Else
                Precio = DBLet(Rs!precioac)
            End If
            If DtoPermitidMor Then Dto = DBLet(Rs!dtoespec)
        ElseIf EsFechaIgualPosterior(Rs!fechanue, FechaOfe, False) Then
        'Precio Nuevo
            If PorCaja Then
                Precio = ComprobarCero(DBLet(Rs!precion1))
            Else
                Precio = ComprobarCero(DBLet(Rs!precionu))
            End If
            If DtoPermitidMor Then Dto = DBLet(Rs!dtoespe1)
        End If
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitidMor = False
        Dto = ""
    End If
    Rs.Close
    Set Rs = Nothing
    ObtenerPrecioEsp = Precio
EObtenerPreEsp:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Especia", Err.Description
End Function


'Private Function ObtenerTarifaArt2(FechaOfe As String, PorCaja As Boolean) As String
''Comprueba si existe una Tarifa de Articulos en la tabla: slista (Tarifas Articulos)
''Dim devuelve As String
'Dim SQL As String
'Dim Rs As ADODB.Recordset
'Dim Precio As String
'
'    On Error GoTo EObtenerTarifa2
'
'    Precio = ""
'
'    'Obtener Precio Actual
'    SQL = "SELECT * FROM slista "
'    SQL = SQL & " WHERE codartic=" & DBSet(CodigoArtic, "T") & " AND codlista=" & CodigoLista
'    'SQL = SQL & " and ( isnull(fechanue) or (fechanue<= '" & Format(FechaOfe, FormatoFecha) & "'))"
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not Rs.EOF Then
'    'Existe una promocion, y salimos con el precio actual o nuevo
'        DtoPermitidMor = CBool(Rs!dtopermi)
'        If IsNull(Rs!fechanue) Or (EsFechaIgualPosterior(FechaOfe, DBLet(Rs!fechanue), False)) Then
'        'Precio Actual
'            If PorCaja Then
'                Precio = ComprobarCero(DBLet(Rs!precioa1))
'            Else
'                Precio = DBLet(Rs!precioac)
'            End If
'        ElseIf EsFechaIgualPosterior(Rs!fechanue, FechaOfe, False) Then
'        'Precio Nuevo
'            If PorCaja Then
'                Precio = ComprobarCero(DBLet(Rs!precion1))
'            Else
'                Precio = ComprobarCero(DBLet(Rs!precionu))
'            End If
'        End If
'    Else  'No existe Precio Especial salir y pasar al siguiente paso
'        Precio = ""
'        DtoPermitidMor = False
'    End If
'    Rs.Close
'    Set Rs = Nothing
'    ObtenerTarifaArt2 = Precio
'EObtenerTarifa2:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Tarifa Artículo", Err.Description
'End Function


'***********************************************************************+
'   Aceites Morales
'   la tarifa es realmente una TO con unas fechas inicio y fin asignadas a un codigo de tarifa
Private Function ObtenerTarifaArt(FechaOfe As String, PorUnidad As Boolean) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerTarifaArt

    ObtenerTarifaArt = ""
    Precio = ""
    SQL = "select pfvu from ariges" & EmprMorales & ".olitarifaoferta,ariges" & EmprMorales & ".olitarifaofertalin where"
    SQL = SQL & " olitarifaoferta.codigo=olitarifaofertalin.codigo and"
    SQL = SQL & " fechaini<='" & Format(FechaOfe, FormatoFecha) & "' and fechafin >='" & Format(FechaOfe, FormatoFecha) & "'"
    SQL = SQL & " and codartic = " & DBSet(CodigoArticMor, "T")
    SQL = SQL & " and Tarifa = " & mvarCodListMor
    SQL = SQL & " order by olitarifaoferta.codigo desc"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        Precio = ComprobarCero(DBLet(Rs!pfvu))

    Else 'No existe
        Precio = ""

    End If
    Rs.Close
    
    ObtenerTarifaArt = Precio

EObtenerTarifaArt:
    If Err.Number <> 0 Then MuestraError Err.Number, "ObtenerTarifaArt"
    Set Rs = Nothing
End Function













Private Function ObtenerPrecioVenta() As String
'Obtiene el precio de Venta al Publico (preciove) de la tabla sartic
    ObtenerPrecioVenta = DevuelveDesdeBDNew(conAri, "ariges" & EmprMorales & ".sartic", "preciove", "codartic", CodigoArticMor, "T")
End Function


Public Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Long
Dim NumCajas As Long
Dim Cantidad As Long, UniPorCaja As Integer
On Error Resume Next

    Cantidad = CLng(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(Cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function



'Abril 2009
'Comprobacion del articulo en las Tarifas-Ofertas (TOs)
Private Function ObtenerTO(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una promocion en la tabla: spromo (Promociones Tarifas)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerTO


    
    'Precio
    SQL = "select pfvu from ariges" & EmprMorales & ".olitarifaoferta,ariges" & EmprMorales & ".olitarifaofertalin where"
    SQL = SQL & " olitarifaoferta.codigo=olitarifaofertalin.codigo and"
    SQL = SQL & " fechaini<='" & Format(FechaOfe, FormatoFecha) & "' and fechafin >='" & Format(FechaOfe, FormatoFecha) & "'"
    SQL = SQL & " and codartic = " & DBSet(CodigoArticMor, "T")
    SQL = SQL & " and codclien = " & mvarCodClieMor
    SQL = SQL & " order by olitarifaoferta.codigo desc"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        Precio = ComprobarCero(DBLet(Rs!pfvu))
        DtoPermitidMor = True
    Else 'No existe
        Precio = ""
        DtoPermitidMor = False
    End If
    Rs.Close
    Set Rs = Nothing
    ObtenerTO = Precio
    
EObtenerTO:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio TOs", Err.Description
End Function



