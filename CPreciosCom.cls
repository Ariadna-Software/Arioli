VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreciosCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Modulo para Calcular el Precio a Aplicar a un PEDIDO de Compras (a PROVEEDOR)
' Y los descuentos correspondientes si se aplica descuento
'TABLAS: slispr, sdtomp

'------------------------------------------------
'Variables del modulo CPreciosCom
'------------------------------------------------

'TABLA: slispr
'---------------
Private mvarCodArtic As String
Private mvarCodProve As String

Private mvarCantFija As Integer 'Cantidad fija
Private mvarCantMin As Integer  'Cantidad minima

Private mvarDtoPermitido As Boolean 'si permite descuento, buscará posibles descuentos en la tabla "sdtomp"

'Datos Promocion
Private mvarFecIniPromo As String
Private mvarFecFinPromo As String
Private mvarPrecioPromo As Currency
Private mvarDtoPermiPromo As Boolean

'Valores actuales
Private mvarPrecioAct As Currency
Private mvarPrecioNue As Currency
Private mvarFechaNue As String

'**********************************
'Laura: campos añadidos el 11/09/06
Private mvarDtoAct1 As Currency
Private mvarDtoAct2 As Currency
'**********************************

'TABLA: sdtomp
'---------------
Private mvarCodFamia As String
Private mvarCodMarca As String



'valor del descuento que devuelve para aplicarlo
Private mvarDescuento1 As String
Private mvardescuento2 As String


'------------------------------------------------
'Propiedades del modulo CPreciosCom
'------------------------------------------------


Public Property Let CantFija(ByVal vData As Integer)
     mvarCantFija = vData
End Property

Public Property Get CantFija() As Integer
     CantFija = mvarCantFija
End Property


Public Property Let CantMinima(ByVal vData As Integer)
     mvarCantMin = vData
End Property

Public Property Get CantMinima() As Integer
     CantMinima = mvarCantMin
End Property


Public Property Let DtoPermitido(ByVal vData As Boolean)
     mvarDtoPermitido = vData
End Property

Public Property Get DtoPermitido() As Boolean
     DtoPermitido = mvarDtoPermitido
End Property

Public Property Let CodigoArtic(ByVal vData As String)
     mvarCodArtic = vData
End Property

Public Property Get CodigoArtic() As String
     CodigoArtic = mvarCodArtic
End Property


Public Property Let CodigoProve(ByVal vData As String)
     mvarCodProve = vData
End Property

Public Property Get CodigoProve() As String
     CodigoProve = mvarCodProve
End Property

Public Property Let CodigoFamia(ByVal vData As String)
     mvarCodFamia = vData
End Property

Public Property Get CodigoFamia() As String
     CodigoFamia = mvarCodFamia
End Property

Public Property Let CodigoMarca(ByVal vData As String)
     mvarCodMarca = vData
End Property

Public Property Get CodigoMarca() As String
     CodigoMarca = mvarCodMarca
End Property


Public Property Let FechaIniPromo(ByVal vData As String)
     mvarFecIniPromo = vData
End Property

Public Property Get FechaIniPromo() As String
     FechaIniPromo = mvarFecIniPromo
End Property


Public Property Let FechaFinPromo(ByVal vData As String)
     mvarFecFinPromo = vData
End Property

Public Property Get FechaFinPromo() As String
     FechaFinPromo = mvarFecFinPromo
End Property


Public Property Let PrecioPromo(ByVal vData As Currency)
     mvarPrecioPromo = vData
End Property

Public Property Get PrecioPromo() As Currency
     PrecioPromo = mvarPrecioPromo
End Property


Public Property Let DtoPermiPromo(ByVal vData As Boolean)
     mvarDtoPermiPromo = vData
End Property

Public Property Get DtoPermiPromo() As Boolean
     DtoPermiPromo = mvarDtoPermiPromo
End Property


Public Property Let PrecioActual(ByVal vData As Currency)
     mvarPrecioAct = vData
End Property

Public Property Get PrecioActual() As Currency
     PrecioActual = mvarPrecioAct
End Property


Public Property Let PrecioNuevo(ByVal vData As Currency)
     mvarPrecioNue = vData
End Property

Public Property Get PrecioNuevo() As Currency
     PrecioNuevo = mvarPrecioNue
End Property

Public Property Let FechaNuevo(ByVal vData As String)
     mvarFechaNue = vData
End Property

Public Property Get FechaNuevo() As String
     FechaNuevo = mvarFechaNue
End Property


'Descuento actual 1: de la tabla slispr
Public Property Let DescuentoActual1(ByVal vData As Currency)
     mvarDtoAct1 = vData
End Property

Public Property Get DescuentoActual1() As Currency
     DescuentoActual1 = mvarDtoAct1
End Property



'Descuento actual 2: de la tabla slispr
Public Property Let DescuentoActual2(ByVal vData As Currency)
     mvarDtoAct2 = vData
End Property

Public Property Get DescuentoActual2() As Currency
     DescuentoActual2 = mvarDtoAct2
End Property






'Descuentos de la tabla sdtomp
Public Property Let Descuento1(ByVal vData As String)
     mvarDescuento1 = vData
End Property

Public Property Get Descuento1() As String
     Descuento1 = mvarDescuento1
End Property


Public Property Let Descuento2(ByVal vData As String)
     mvardescuento2 = vData
End Property

Public Property Get Descuento2() As String
     Descuento2 = mvardescuento2
End Property


'------------------------------------------------
'Metodos del modulo CPreciosCom
'------------------------------------------------

Public Function Leer(codArtic As String, codProve As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
On Error GoTo ELeer

    SQL = "SELECT * FROM slispr "
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codprove=" & codProve
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        mvarCantFija = DBLet(Rs!CantFija, "N")
        mvarCantMin = DBLet(Rs!cantmini, "N")

        mvarDtoPermitido = Rs!dtopermi
        mvarCodArtic = DBLet(Rs!codArtic, "T")
        mvarCodProve = Rs!codProve
        
        'datos de promocion
        mvarFecIniPromo = DBLet(Rs!FechaIni, "F")
        mvarFecFinPromo = DBLet(Rs!FechaFin, "F")
        mvarPrecioPromo = DBLet(Rs!preciopr, "N")
        mvarDtoPermiPromo = DBLet(Rs!dtoperm1, "B")
        
        'valores actuales
        mvarPrecioAct = DBLet(Rs!precioac, "N")
        mvarPrecioNue = DBLet(Rs!precionu, "N")
        mvarDtoAct1 = DBLet(Rs!dtoline1, "N")
        mvarDtoAct2 = DBLet(Rs!dtoline2, "N")
        
        
        'Octubre 2009
        '---------------------------------
        'Si no han puesto fecha nue significa que para
        'ese articulo y proveedor tiene que coger el valo de actual
        'COn lo cual establezco la fechanue(si es nulo) al año 2100, asi siempre coge valor actual
        'ANTES: mvarFechaNue = DBLet(RS!fechanue, "F")
        If IsNull(Rs!fechanue) Then
            mvarFechaNue = CDate("01/01/2100")
        Else
            mvarFechaNue = Rs!fechanue
        End If
        
        
        Leer = True
    End If
    Rs.Close
    Set Rs = Nothing
    
ELeer:
    If Err.Number <> 0 Then Leer = False
End Function



Public Function ComprobarCantidad2(Cantidad As Long) As Boolean
'Comprueba si la cantidad a insertar en la linea de pedido o Albaran de Compras
'cumple la cantidad fija o minima si alguna de las 2 tiene valor para el artic-prove

    ComprobarCantidad2 = False
    
    'Si hay cantidad Fija solo se pueden comprar multiplos de esa cantidad
    If CantFija > 0 Then
        'comprobar q la cantidad es multiplo de CantFija
        If Not (Cantidad >= CantFija) Or Not (Cantidad Mod CantFija = 0) Then
            MsgBox "Cantidad Fija a comprar del artículo " & CodigoArtic & " es:  " & CantFija & "  (o múltiplos)", vbExclamation
            Exit Function
        End If
        
    'Si hay cantidad minima no se puede comprar menos de esa cantidad
    ElseIf CantMinima > 0 Then
        If Not (Cantidad >= CantMinima) Then
            MsgBox "Cantidad Mínima a comprar del artículo " & CodigoArtic & " es:  " & CantMinima, vbExclamation
            Exit Function
        End If
    End If
    
    ComprobarCantidad2 = True
End Function



Public Function ObtenerPrecio(FechaPed As String) As String
'FechaPed -> IN: Fecha del Pedido para comprobar si esta en los periodos de promocion, etc...
'RETURN: valor del precio
Dim PrecioCom As String
Dim Dto1Com As String
Dim Dto2Com As String

    On Error GoTo EObtenerPre

   
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla "slispr" si encuentro valor salgo con ese precio
    PrecioCom = ObtenerPromocion(FechaPed)
    If PrecioCom <> "" Then 'Ha obtenido el Precio
        If DtoPermiPromo Then
            'mirar si los campos de descuento de la tabla slispr tienen valor
            If (DescuentoActual1 + DescuentoActual2) > 0 Then
                Dto1Com = DescuentoActual1
                Dto2Com = DescuentoActual2
            Else
                'Buscar en la tabla "sdtomp" los descuentos
                Dto1Com = ObtenerDescuentos(FechaPed, Dto2Com)
            End If
        End If
        
    Else 'No encuentra Precio de Promocion en slispr y Buscar en la misma tabla: slispr (Tabla Cab. Precios Prov.)
        'COMPROBAR PRECIO Prov.
        '===========================================================================
        PrecioCom = ObtenerPrecioProv(FechaPed)
        If PrecioCom <> "" Then 'Ha obtenido el precio
            'mirar si los campos de descuento de la tabla slispr tienen valor
            If (DescuentoActual1 + DescuentoActual2) > 0 Then
                Dto1Com = DescuentoActual1
                Dto2Com = DescuentoActual2
            ElseIf DtoPermitido Then
            'Buscar en la tabla sdtomp los descuentos
                Dto1Com = ObtenerDescuentos(FechaPed, Dto2Com)
            End If
        Else 'Buscar en sartic el precio ultima compra del articulo (preciouc)
            PrecioCom = ObtenerPrecioUltCompra
        End If
    End If
    
    ObtenerPrecio = ComprobarCero(PrecioCom)
    Descuento1 = ComprobarCero(Dto1Com)
    Descuento2 = ComprobarCero(Dto2Com)
    
EObtenerPre:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Precio", Err.Description
End Function


Private Function ObtenerPromocion(FechaPed As String) As String
'Comprueba si existe una promocion en la tabla: slispr (Cab. Precios Prov.)
Dim Precio As String
On Error GoTo EObtenerPromo

    Precio = ""
    
    'Existe una promocion para ese articulo y proveedor, comprobar fechas
    If FechaIniPromo <> "" And FechaFinPromo <> "" Then
        If EntreFechas(FechaIniPromo, FechaPed, FechaFinPromo) Then
            'Dentro Periodo Promocion Actual
            Precio = PrecioPromo
        End If
    Else 'No existe Promocion salir y pasar al siguiente paso
        Precio = ""
        DtoPermiPromo = False
    End If
    ObtenerPromocion = Precio
    
EObtenerPromo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Promoción", Err.Description
End Function


Private Function ObtenerDescuentos(FechaPed As String, Dto2Com) As String
'IN -> FechaPed: fecha del pedido para comprobar si hay algun descuento valido en esa fecha
'                para el Proveedor, y para la familia y marca del Articulo
'OUT -> ObtenerDescuentos: Aqui se devuelve el valor del Dto1 si lo encuentra
'        Dto2com: Devuelve el valor del Dto2 si lo encuentra
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String
    
    'Obtenemos la Familia y la Marca del Articulo
    vCodMarca = "Codmarca"
    CodigoFamia = DevuelveDesdeBDNew(conAri, "sartic", "Codfamia", "codartic", CodigoArtic, "T", vCodMarca)
    CodigoMarca = vCodMarca
    
    CampoDto1 = "dtoline1"
    campoDto2 = "dtoline2"
    
    'Obtener Descuento
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & " FROM sdtomp "
    SQL = SQL & " WHERE codprove=" & CodigoProve & " AND codfamia=" & CodigoFamia & " AND codmarca=" & CodigoMarca
    SQL = SQL & " and (fechadto<= '" & Format(FechaPed, FormatoFecha) & "')"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentos = DBLet(Rs.Fields(0).Value, "N")
        Dto2Com = DBLet(Rs.Fields(1).Value, "N")
    Else
        ObtenerDescuentos = ""
        Dto2Com = ""
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Function ObtenerPrecioProv(FechaPed As String) As String
'Comprueba si existe un Precio de Proveedor en la tabla: slispr (Cab. Precios Prov.)
Dim Precio As String

    On Error GoTo EObtenerPreProv

    Precio = ""
    
    If FechaNuevo <> "" Then
    'Existe un precio para el Proveedor, y salimos con el precio actual o nuevo
        
        'Laura: modificado el 11/09/06
        '************************************************************
'        If (EsFechaIgualPosterior(FechaPed, FechaNuevo, False)) Then
'            'Precio Actual
'            Precio = CStr(PrecioActual)
'        ElseIf EsFechaIgualPosterior(FechaNuevo, FechaPed, False) Then
'            'Precio Nuevo
'            Precio = CStr(PrecioNuevo)
'        End If
        If EsFechaIgualPosterior(FechaNuevo, FechaPed, False) Then
            'Precio Nuevo
            Precio = CStr(PrecioNuevo)
        Else
            'Precio Actual
            Precio = CStr(PrecioActual)
        End If
        '*************************************************************
        
    'Laura: añadido el 11/09/06
    '*****************************
    ElseIf PrecioActual > 0 Then
        Precio = PrecioActual
    '*****************************
    
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    ObtenerPrecioProv = Precio
    
EObtenerPreProv:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Proveedor", Err.Description
End Function


Private Function ObtenerPrecioUltCompra() As String
'Obtiene el precio de ultima compra (preciouc) de la tabla sartic
    ObtenerPrecioUltCompra = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", CodigoArtic, "T")
End Function

