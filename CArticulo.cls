VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ARTICULOS
'TABLA: sartic

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodartic As String 'Cod. articulo
Private mNomArtic As String 'Nombre articulo

Private mCodFamia As String 'Familia del Articulo
Private mCodCateg As String 'categoria del Articulo

Private mCodStatu As Byte 'cod. estatus: 1=Bloqueado, 2=Caducado
Private mCtrStock As Byte 'hay control de stock (si/No)
Private mArtVario As Byte 'Si es articulo de varios o no
Private mConjunto As Boolean  'Si es articulo de conjunto o no

Private mCodigIVA As Byte

Private mPrecioVe As String 'Precio de venta (P.V.P.)
Private mUniCajas As Integer 'unidades por caja
Private mPreciomp As Currency 'Precio medio ponderado
Private mPreciost As Currency 'Precio standard
Private mPreciouc As Currency 'Precio ultima compra
Private mMargeCom As Currency 'margen comercial

Private mTextoVen As String 'texto a mostrar para ventas
Private mTextoCom As String 'texto a mostrar para compras

Private mCodProve As Long   'codigo proveedor

Private mFactorConversion As Currency   'Factor de conversion arioli
Private mLitrosxUd As Currency

Private mTrazabilidad As Boolean

'------------------------------------------------
'Propiedades del modulo CArticulo
'------------------------------------------------

'**** Codigo del articulo
Public Property Let Codigo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodartic
End Property


'**** Nombre del Articulo
Public Property Let Nombre(ByVal vData As String)
     mNomArtic = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomArtic
End Property



'**** Familia del Articulo
Public Property Let Familia(ByVal vData As String)
     mCodFamia = vData
End Property

Public Property Get Familia() As String
     Familia = mCodFamia
End Property


'**** Categoria del Articulo
Public Property Let Categoria(ByVal vData As String)
     mCodCateg = vData
End Property

Public Property Get Categoria() As String
     Categoria = mCodCateg
End Property



'**** cod. status del Articulo
Public Property Let Status(ByVal vData As Byte)
     mCodStatu = vData
End Property

Public Property Get Status() As Byte
     Status = mCodStatu
End Property



'**** Articulo tiene control de stock (si/no)
Public Property Let CtrStock(ByVal vData As Byte)
     mCtrStock = vData
End Property

Public Property Get CtrStock() As Byte
     CtrStock = mCtrStock
End Property



'**** Articulo es de Varios (Si/No)
Public Property Let EsDeVarios(ByVal vData As Byte)
     mArtVario = vData
End Property

Public Property Get EsDeVarios() As Byte
     EsDeVarios = mArtVario
End Property



'**** Articulo es de Conjunto (Si/No)
Public Property Let EsConjunto(ByVal vData As Boolean)
     mConjunto = vData
End Property

Public Property Get EsConjunto() As Boolean
     EsConjunto = mConjunto
End Property



'**** Precio de Venta
Public Property Let PrecioVenta(ByVal vData As String)
     mPrecioVe = vData
End Property

Public Property Get PrecioVenta() As String
     PrecioVenta = mPrecioVe
End Property


'**** Unidades por caja
Public Property Let UnidCaja(ByVal vData As Integer)
     mUniCajas = vData
End Property

Public Property Get UnidCaja() As Integer
     UnidCaja = mUniCajas
End Property



'**** Precio medio ponderado
Public Property Let PrecioMedPon(ByVal vData As Currency)
     mPreciomp = vData
End Property

Public Property Get PrecioMedPon() As Currency
     PrecioMedPon = mPreciomp
End Property


'**** Precio standard
Public Property Let PrecioStan(ByVal vData As Currency)
     mPreciost = vData
End Property

Public Property Get PrecioStan() As Currency
     PrecioStan = mPreciost
End Property



'**** Precio ultima compra
Public Property Let PrecioUltCom(ByVal vData As Currency)
     mPreciouc = vData
End Property

Public Property Get PrecioUltCom() As Currency
     PrecioUltCom = mPreciouc
End Property


'**** Margen comercial
Public Property Let MargenComercial(ByVal vData As Currency)
     mMargeCom = vData
End Property

Public Property Get MargenComercial() As Currency
     MargenComercial = mMargeCom
End Property



'**** Tipo de IVA del articulo
Public Property Let TipoIVA(ByVal vData As Byte)
     mCodigIVA = vData
End Property

Public Property Get TipoIVA() As Byte
     TipoIVA = mCodigIVA
End Property




'**** Texto para Ventas
Public Property Let TextoVentas(ByVal vData As String)
     mTextoVen = vData
End Property

Public Property Get TextoVentas() As String
     TextoVentas = mTextoVen
End Property


'**** Texto para Compras
Public Property Let TextoCompras(ByVal vData As String)
     mTextoCom = vData
End Property

Public Property Get TextoCompras() As String
     TextoCompras = mTextoCom
End Property


'**** Codigo proveedor
Public Property Let codProve(ByVal vData As Long)
     mCodProve = vData
End Property

Public Property Get codProve() As Long
     codProve = mCodProve
End Property


'**** Factor conversion. SOLO TIENE GET

Public Property Get FactorConversion() As Currency
     FactorConversion = mFactorConversion
End Property


'**** Litros x Unidad. SOLO TIENE GET

Public Property Get LitrosxUd() As Currency
     LitrosxUd = mLitrosxUd
End Property








'*****   Trazabilidad. Solo tiene GET
Public Property Get Trazabilidad() As Boolean
     Trazabilidad = mTrazabilidad
End Property







'------------------------------------------------
'Procedimientos del modulo CArticulo
'------------------------------------------------

Public Function Existe(vCodArtic As String) As Boolean
'Comprueba si existe el articulo en la BD
Dim devuelve As String
Dim Nom As String

    On Error GoTo EExiste
    
    Nom = "nomartic"
    devuelve = DevuelveDesdeBD(conAri, "codartic", "sartic", "codartic", vCodArtic, "T", Nom)
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Art�culo " & vCodArtic & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
        Codigo = vCodArtic
        Nombre = Nom
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodArtic As String) As Boolean
'Leer los datos de un Articulo dado
'Lee de la BD: Ariges, Tabla: sartic
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codartic,nomartic,codfamia,codstatu,ctrstock,preciove,unicajas,codigiva,artvario,conjunto,textoven,textocom,codcateg,preciomp,preciost,preciouc,margecom,codprove,FactorConversion,LitrosUnidad,Trazabilidad "
    SQL = SQL & " FROM sartic "
    SQL = SQL & " WHERE codartic=" & DBSet(vCodArtic, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        Codigo = CStr(RS!CodArtic)
        
        Nombre = RS!NomArtic
        Familia = RS!Codfamia
        Categoria = DBLet(RS!codCateg, "T")
        
        Status = RS!codstatu
        CtrStock = RS!CtrStock
        EsDeVarios = RS!artvario
        Me.EsConjunto = RS!Conjunto
        
        PrecioVenta = RS!preciove
        TipoIVA = RS!codigiva
        UnidCaja = RS!UniCajas
        
        PrecioMedPon = DBLet(RS!precioMP, "N") 'precio medio ponderado
        PrecioStan = DBLet(RS!preciost, "N") 'precio standard
        PrecioUltCom = DBLet(RS!PrecioUC, "N") 'precio ultima compra
        MargenComercial = DBLet(RS!margecom, "N") 'margen comercial
        
        TextoVentas = DBLet(RS!textoven, "T")
        TextoCompras = DBLet(RS!textocom, "T")
        
        codProve = DBLet(RS!codProve, "N")
        
        'NO puede ser cero
        mFactorConversion = DBLet(RS!FactorConversion, "N")
        If mFactorConversion = 0 Then mFactorConversion = 1
        
        'Listros por unidad
        mLitrosxUd = DBLet(RS!LitrosUnidad, "N")
        If mLitrosxUd = 0 Then mLitrosxUd = 1 'Para que no de errores
        
        mTrazabilidad = DBLet(RS!Trazabilidad, "N") = 1
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Art�culo" & vbCrLf
    SQL = SQL & "N�mero: " & Err.Number & vbCrLf
    SQL = SQL & "Descripci�n: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Sub MostrarStatusArtic(Bloquea As Boolean)
'Recuperamos el estatus (codstatu) del articulo, y si tiene
'valor 1 o 2 lo mostramos en un mensaje
'codstatu=0 normal
'codstatu=1 bloqueado, muestra un mensaje
'codstatu=2 caducado, muestra un mensaje y bloquea linea de articulo

    Bloquea = False

    Select Case Status
        ' 0 = Normal
        Case 1 'BLOQUEADO
            MsgBox "El articulo " & Codigo & " esta BLOQUEADO." & vbCrLf & "No se puede insertar una linea con este art�culo.", vbExclamation, "Situaci�n art�culo"
            Bloquea = True
        Case 2 'CADUCADO
            MsgBox "El articulo " & Codigo & " esta CADUCADO.", vbInformation, "Situaci�n art�culo"
    End Select
End Sub


Public Function ExisteEnAlmacen(codAlm As String, Optional CanStock As String) As Boolean
Dim devuelve As String
Dim Cantidad As String

    On Error GoTo EExisteAlm

    'Comprobar que existe de ese articulo en el almacen seleccionado
    Cantidad = "canstock"
    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Codigo, "T", Cantidad, "codalmac", codAlm, "N")
    
    If devuelve = "" Then
        MsgBox "No existe unidades del Art�culo: " & Codigo & "  en el Almacen: " & codAlm, vbExclamation
        ExisteEnAlmacen = False
    Else
        ExisteEnAlmacen = True
        CanStock = Cantidad
    End If
    
EExisteAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Art�culo", Err.Description
End Function




Public Function ExistenciaTotalAlmacenes() As Currency
'-- devuelve la cantidad de stock total en todos los almacenes del articulo
Dim rst As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo EExisTotal
    
    If Codigo <> "" Then
        SQL = "SELECT SUM(canstock) FROM salmac where codartic=" & DBSet(Codigo, "T")
        Set rst = New ADODB.Recordset
        rst.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            ExistenciaTotalAlmacenes = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    Exit Function
    
EExisTotal:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener existencia total en Almacenes del Art�culo", Err.Description
End Function


Public Function EnInventario(codAlm As String) As Boolean
'SI  se esta haciendo inventario del articulo o no
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "statusin", "codartic", Codigo, "T", , "codalmac", codAlm, "N")
    
    If devuelve = "1" Then
        EnInventario = True
        MsgBox "El art�culo " & Codigo & " esta inventariandose.", vbInformation
    Else
        EnInventario = False
    End If
End Function


Public Function EsInstalacion() As Boolean
'Coprobar si el articulo es instalacion
Dim devuelve As String
    
    If Familia = "" Then _
        Familia = DevuelveDesdeBDNew(conAri, "sartic", "codfamia", "codartic", Codigo, "T")
        
    
    devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", Familia, "N")
    If devuelve = "0" Then 'No es instalacion
        EsInstalacion = False
    ElseIf devuelve = "1" Then
        EsInstalacion = True
    End If

End Function


Public Sub MostrarTextoCom()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoCompras, vbInformation, "Texto para compras"
End Sub


Public Sub MostrarTextoVen()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoVentas, vbInformation, "Texto para ventas"
End Sub



Public Function ActualizarUltFechaCompra_(vFecha As String, vPrecio As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
Dim SQL As String
Dim actualiza As Boolean

    On Error GoTo EActFecha

    'Obtenemos la ult. fecha de compra que tiene el articulo
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "ultfecco", "codartic", Codigo, "T")
    
    actualiza = False
    
    If SQL = "" Then
        'No tiene fecha ult. compra
        actualiza = True
    ElseIf CDate(vFecha) >= CDate(SQL) Then
        'fecha de compra posterior
        actualiza = True
    End If
    
    If actualiza Then
        SQL = "UPDATE sartic SET ultfecco=" & DBSet(vFecha, "F")
        'actualizar el ult. precio de compra si el nuevo valor no es 0
        If CCur(vPrecio) > 0 Then SQL = SQL & ",preciouc=" & DBSet(vPrecio, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
        Conn.Execute SQL
    Else
        SQL = "Fecha de compra anterior a una compra realizada." & vbCrLf & "Fecha ultima compra: " & SQL
        SQL = SQL & vbCrLf & vbCrLf & Nombre
        MsgBox SQL, vbExclamation
    End If
    
EActFecha:
    If Err.Number <> 0 Then
        ActualizarUltFechaCompra_ = False
    Else
        ActualizarUltFechaCompra_ = True
    End If
End Function



Public Function ActualizarPrecioMedPond(vCantidad As Currency, vPrecio As Currency, Optional vCantAnt As Currency) As Boolean
'-- Modificar precio medio ponderado en la tabla sartic, se utiliza al crear
'-- albaranes de compra
'(IN) vCantidad=cantidad comprada o modificada
'(IN) vPrecio= precio de compra
'(IN) vCantAnt= cantidad anterior a la modificacion
Dim SQL As String
Dim actualiza As Boolean
Dim existencia As Currency
Dim Importe As Currency

    On Error GoTo EActPMP

    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", Codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    actualiza = False
    If PrecioMedPon = 0 Then
        'se actualiza con el precio de compra
        PrecioMedPon = vPrecio
        actualiza = True
    Else
        'se actualiza con la formula:
        '((existencia_total_almacenes_antes_compra * precio_medio_ponderado_actual) + importe_compra) / (existencia_total_almacenes_antes_compra + cantidad_compra)
        
        existencia = ExistenciaTotalAlmacenes
        'como aqui ya le habiamos sumado la cantidad q acabamos de comprar
        'se la quitamos para obtener la existencia antes de la compra
'        existencia = existencia - vCantidad

        'si hemos modificado la linea del albaran, a la existencia habr� q
        'quitarle la cantidad q habia antes
        If vCantAnt <> 0 Then existencia = existencia - vCantAnt
        
        'obtenemos el nuevo precio_medio_ponderado a partir del valor actual
        If (existencia + vCantidad) <> 0 Then
            Importe = Round2((vCantidad * vPrecio), 2)
            PrecioMedPon = Round2(((existencia * PrecioMedPon) + Importe) / (existencia + vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
            actualiza = True
        End If
    End If

    
    If actualiza Then
        SQL = "UPDATE sartic SET preciomp=" & DBSet(PrecioMedPon, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
        Conn.Execute SQL
    End If
    
    ActualizarPrecioMedPond = True
    Exit Function
    
EActPMP:
'    If Err.Number <> 0 Then
        ActualizarPrecioMedPond = False
'    Else
'        ActualizarPrecioMedPond = True
'    End If
End Function






Public Function ReestablecerUltFechaCompra(vAlmac As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
'con los valores del ultimo movimiento que encontremos en la smoval
Dim SQL As String
Dim cadSel As String
Dim RS As ADODB.Recordset

    On Error GoTo ERestFecha

    'Obtenemos el ultimo movimiento de albaran de compra que tiene el articulo
    cadSel = " detamovi='ALC' and codartic=" & DBSet(Codigo, "T") & " AND codalmac=" & vAlmac & " AND and tipomovi=1 "
    
    SQL = "select * from smoval WHERE " & cadSel
    SQL = SQL & " and horamovi =(select max(horamovi) from smoval where " & cadSel & ")"

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'actualizamos la sartic con esos valores
        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!Fechamov, "F")
        SQL = SQL & ", preciouc=" & DBSet(RS!impormov, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T")
        Conn.Execute SQL
    End If
    RS.Close
    Set RS = Nothing

    
ERestFecha:
    If Err.Number <> 0 Then
        ReestablecerUltFechaCompra = False
    Else
        ReestablecerUltFechaCompra = True
    End If
End Function




Public Function ReestablecerPrecioMedPon(vCantidad As Currency, vPrecio As Currency) As Boolean
'reestablecer el valor del precio medio ponderado
'       (pmp_actual * stock_actual) - (vCantidad * vPrecio)
'pmp=   ---------------------------------------------------
'       (stock_actual - cantidad)

Dim existencia As Currency 'existencia de stock en almacenes
Dim SQL As String
Dim Importe As Currency

    On Error GoTo EResPMP
    
    ReestablecerPrecioMedPon = False
    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", Codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    
    existencia = ExistenciaTotalAlmacenes
    
    If (existencia - vCantidad) <> 0 Then
        Importe = Round2((vCantidad * vPrecio), 2)
        PrecioMedPon = Round2(((PrecioMedPon * existencia) - Importe) / (existencia - vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
    Else
        PrecioMedPon = 0
    End If
    
    SQL = "UPDATE sartic SET preciomp=" & DBSet(PrecioMedPon, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
    Conn.Execute SQL
    
    ReestablecerPrecioMedPon = True
    Exit Function
    
EResPMP:
    ReestablecerPrecioMedPon = False
    MsgBox Err.Number, "Reestablecer precio medio ponderado.", Err.Description
End Function




Public Function ObtenerPorceIVA() As Currency
'Devuelve el % de IVA del articulo correspondiente al codigiva
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo EPorce
    
    If Codigo <> "" Then 'codartic
        'Obtenemos el % de IVA a aplicar
        SQL = "SELECT porceiva FROM tiposiva WHERE codigiva=" & TipoIVA
        Set RS = New ADODB.Recordset
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL = ""
        If Not RS.EOF Then SQL = RS.Fields(0).Value
        
        RS.Close
        Set RS = Nothing
        ObtenerPorceIVA = CCur(SQL)
    End If
    
EPorce:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Obteniendo porcentaje de IVA." & vbCrLf
        SQL = SQL & "N�mero: " & Err.Number & vbCrLf
        SQL = SQL & "Descripci�n: " & Err.Description
        MsgBox SQL, vbExclamation
        Set RS = Nothing
    End If
End Function


''David
''Modifica:  Para que guarde en una variable el valor del % del iva
'Public Function ObtenerPrecioConIVA2(ByRef CadenaIVA As String, Optional newPrecio As String) As Currency
''Devuelve el precio del articulo tras aplicarle el IVA correspondiente
''si newPrecio tiene valor le aplicamos el IVA a ese precio ya que habremos
''introducido el precio del articulo manualmente
''si no le pasamos newPrecio se aplicar� al precio_venta del articulo
'Dim SQL As String
'Dim Porce As Currency ' % de IVA
'
'    On Error GoTo EPrecio
'
'    If Codigo <> "" Then 'codartic
'        'Obtener el % de IvA q se le aplica al articulo
'        Porce = ObtenerPorceIVA
'        CadenaIVA = CStr(Porce)
'        If Porce <> CCur(0) Then
'            If newPrecio = "" Then
'                'al precio de venta del articulo
'                SQL = CalcularDto(PrecioVenta, CStr(Porce))
'                ObtenerPrecioConIVA2 = Round(CCur(ComprobarCero(PrecioVenta)) + CCur(ComprobarCero(SQL)), 4)
'            Else
'                'al precio especial del articulo (para un cliente tn promociones)
'                SQL = CalcularDto(newPrecio, CStr(Porce))
'                ObtenerPrecioConIVA2 = Round(CCur(ComprobarCero(newPrecio)) + CCur(ComprobarCero(SQL)), 4)
'            End If
'        Else
'            If newPrecio = "" Then
'                ObtenerPrecioConIVA2 = CCur(ComprobarCero(PrecioVenta))
'            Else
'                ObtenerPrecioConIVA2 = CCur(ComprobarCero(newPrecio))
'            End If
'        End If
'    End If
'
'EPrecio:
'    If Err.Number <> 0 Then
'        SQL = "Se ha producido un error. " & "Calculando precio con IVA." & vbCrLf
'        SQL = SQL & "N�mero: " & Err.Number & vbCrLf
'        SQL = SQL & "Descripci�n: " & Err.Description
'        MsgBox SQL, vbExclamation
'    End If
'End Function









Public Function ObtenerPrecioSinIVA(Optional newPrecio As String) As Currency
'Devuelve el precio del articulo tras quitarle el IVA correspondiente
'IN -> newPrecio: precio con IVA
'si newPrecio tiene valor le quitamos el IVA a ese precio ya que habremos
'introducido el precio del articulo manualmente
'si no le pasamos newPrecio ser� al precio_venta del articulo
Dim SQL As String
Dim Porce As Currency ' % de IVA
Dim PreuSinIVA As Currency

    On Error GoTo EPrecio
    
    If Codigo <> "" Then 'codartic
        'Obtener el % de IvA q se le aplica al articulo
        Porce = ObtenerPorceIVA
        Porce = Porce / 100
        
        'precio al que le vamos a quitar el IVA
        If newPrecio = "" Then
            'al precio de venta del articulo
            PreuSinIVA = CCur(ComprobarCero(PrecioVenta))
        Else
            'al precio especial del articulo (para un cliente tn promociones)
            'o art. de varios y precio manual
            PreuSinIVA = CCur(ComprobarCero(newPrecio))
            If Porce <> CCur(0) Then
                'quitarle el IVA al precio
                PreuSinIVA = PreuSinIVA / (1 + Porce)
                PreuSinIVA = Round(PreuSinIVA, 4)
            End If
        End If
        ObtenerPrecioSinIVA = PreuSinIVA
    End If

'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
    Exit Function
    
EPrecio:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Calculando precio sin IVA." & vbCrLf
        SQL = SQL & "N�mero: " & Err.Number & vbCrLf
        SQL = SQL & "Descripci�n: " & Err.Description
        MsgBox SQL, vbExclamation
    End If
End Function


'Nuevo###
'Cambiado por David 14 Mayo 2008
' Si tiene descuentos linea, debe calcular el importe con respecto a esos descuentos
Public Function ObtPrecioParaCliente2(codCli As String, canti As String, FecVen As String, des1 As String, des2 As String) As String
'Obtiene el precio que se le va a aplicar a un cliente
'en la venta de un determinado articulo
'Habr� clientes con promociones, precios especiales, etc.

'IN -> codclien: cliente para el q comprobamos si hay precio especial
       'canti: cantidad a vender para saber si se aplica precio por caja
'     fecven: fecha de la venta

'OUT -> des, des2: descuento 1 y descuento 2

Dim cPrecioF As CPreciosFact
Dim NumCajas As Integer
Dim RestoUnid As Integer
Dim PorCaja As Boolean
Dim OrigP As String
Dim Precio As String
Dim cadMen As String
Dim impo As Currency
    On Error GoTo ErrPrecioCli

    'Obtener el precio correspondiente y los descuentos
        
        
    'Dejo los dtos a cero
    des1 = "0"
    des2 = "0"
    'Comprobar si el articulo se vende por cajas antes de entrar a la funci�n
    If UnidCaja >= 0 Then '+ de 1 unidad x caja, significa q se vende por cajas
        Set cPrecioF = New CPreciosFact
        
        'Si se puede vender por cajas(unidcaja>1) poner numero de cajas en una linea con el
        'precio de caja, y otra linea con el resto unidades con precio unidad
'            Cantidad = txtAux(Index).Text
        NumCajas = cPrecioF.ObtenerNumCajas(canti, UnidCaja)
        RestoUnid = CInt(ComprobarCero(canti)) - NumCajas * CInt(UnidCaja)
            
        'Obtenemos la Tarifa del Cliente
        cPrecioF.CodigoLista = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", codCli, "N")
                
        cPrecioF.CodigoArtic = Codigo
        cPrecioF.CodigoClien = codCli
        PorCaja = (NumCajas > 0)
        Precio = cPrecioF.ObtenerPrecio(PorCaja, FecVen, OrigP)
        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
        'Ya que a regresado con pvp del Articulo
        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
            cadMen = "El Art�culo puede venderse por Cajas (" & UnidCaja & "uds. por Caja)." & vbCrLf
            cadMen = cadMen & vbCrLf & "�Desea insertar dos Lineas?:   " & vbCrLf
            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(UnidCaja) & " uds a Precio Caja"
            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(canti) - NumCajas * CInt(UnidCaja) & " uds a Precio Unidad"
            If MsgBox(cadMen, vbInformation + vbYesNo) = vbYes Then
                canti = NumCajas * CInt(UnidCaja)
            Else
                'si no vendemos por cajas traer el precio de la funcion poniendo
                'venta por cajas a false.
                Precio = cPrecioF.ObtenerPrecio(False, FecVen, OrigP)
            End If
         
        Else
            If CCur(cPrecioF.Descuento1) > 0 Or CCur(cPrecioF.Descuento2) > 0 Then
                des1 = cPrecioF.Descuento1
                des2 = cPrecioF.Descuento2
            End If
        End If
        
        ObtPrecioParaCliente2 = Precio
        Set cPrecioF = Nothing
    End If
    Exit Function
    
ErrPrecioCli:
    MuestraError Err.Number, "Obtener precio para el cliente", Err.Description
End Function



Public Function TieneNumLote() As Boolean
'si el articulo pertenece a una categoria que tiene control de numero de lote
Dim cad As String

    If Categoria <> "" Then
        'ver si la categoria a la q pertenece lleva control de lotes
        cad = DevuelveDesdeBDNew(conAri, "scateg", "ctrlotes", "codcateg", Categoria, "T")
        TieneNumLote = (cad = "1")
    Else
        'el articulo no pertenece a una categoria
        TieneNumLote = False
    End If
End Function



Public Function AplicarMargenComercial() As Currency
'Aplica el margen comercial que queremos obtener al precio de ultima compra
'y obtiene el nuevo precio de venta al publico.
Dim margen As Currency

    If Me.MargenComercial <> 0 Then 'el campo margen comercial tiene valor
'        margen = CalcularDto(Me.PrecioUltCom, Me.MargenComercial)
'        margen = Round(margen, 4)
        margen = Round2(((Me.PrecioUltCom * Me.MargenComercial) / 100), 4)
        AplicarMargenComercial = Me.PrecioUltCom + margen 'nuevo pvp
        
'    ElseIf Me.PrecioUltCom > Me.PrecioVenta Then
'        AplicarMargenComercial = Me.PrecioUltCom
        
    Else
        AplicarMargenComercial = Me.PrecioVenta
    End If
    
End Function


Public Function CambiaPrecioVenta(newPrecio As Currency, newMargen As Currency, fecCambio As String, cadError) As Boolean
'actualiza el precio de venta del articulo (PVP) y la fecha en que se realiza la modificacion
'actualiza el campo sartic.preciove
'actualiza el campo sartic.ultfecpvp
Dim SQL As String

    On Error GoTo ErrPrecio

    CambiaPrecioVenta = False
    If Codigo <> "" Then
        SQL = "UPDATE sartic SET preciove=" & DBSet(newPrecio, "N")
        SQL = SQL & ", margecom=" & DBSet(newMargen, "N")
        
        '## LAURA 23/06/2008
        '   actualizar la fecha de ultima modificacion PVP
        SQL = SQL & ", ultfecpvp=" & DBSet(fecCambio, "F")
        '##
        
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T")
        Conn.Execute SQL
        CambiaPrecioVenta = True
    Else
        cadError = cadError & "Falta c�digo del art�culo para actualizar precio."
    End If
    
    Exit Function

ErrPrecio:
    cadError = cadError & "Cambiar precio art�culo." & vbCrLf
    cadError = cadError & Err.Number & ": " & Err.Description
'    MuestraError Err.Number, "Cambiar precio art�culo.", Err.Description
End Function



Private Function TarifasTieneMargen() As Boolean

End Function



'## Laura 22/11/2006
'Public Function ActualizarStockLinConjunto(ByRef cSto As CStock, cadError As String) As Boolean
''para el art�culo q estamos actualizando su stock
'' comprobar si es de conjunto, y si lo es Actualizar stock para cada linea del conjunto
'Dim b As Boolean
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim cStock1 As CStock 'stock de las sublineas de conjunto
'
'    On Error GoTo ErrActConj
'
'    If Me.EsConjunto Then
'        'para cada linea de conjunto del articulo actualizar stock e insertar movimiento
'        SQL = "SELECT * FROM sarti1 WHERE codartic=" & DBSet(Me.Codigo, "T")
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF And b
''            Set cStock1 = New CStock
'            Set cStock1 = cSto
'            cStock1.Cantidad = cSto.Cantidad * DBLet(RS!codarti1, "N")
'            b = cStock1.ActualizarStock
'            'precios!!!!
'
'
'            Set cStock1 = Nothing
'            RS.MoveNext
'        Wend
'        Set RS = Nothing
'    Else
'        'si no es cjto no hacer nada salimo con ok
'        b = True
'    End If
'
'
'    ActualizarStockLinConjunto = b
'    Exit Function
'
'ErrActConj:
'    ActualizarStockLinConjunto = False
'    cadError = Err.Description
'End Function

