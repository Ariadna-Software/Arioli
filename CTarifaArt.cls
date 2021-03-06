VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTarifaArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TARIFAS DE ARTICULOS
'TABLA: slista


'*************************************************************************
'*************************************************************************
'
'       ARIOLI.
'       Cada tarifa tiene entradas en las TOs, con valores inferiores a
'       100000 y, importante, con una referencia en la cabecera de las TO
'       a la tarifa en cuestion.
'       Es decir, en ARIOLI, para saber el precio de un articulo para una
'       tarifa habra que ir a la TO con campo.tarifa = a la del cliente
'       y la FECHA entre las fechaini y fechafin de la to


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodartic As String 'Cod. articulo
Private mCodLista As Integer 'cod. Tarifa -> starif

Private mPrecioac As Currency 'Precio actual
Private mPrecioa1 As Currency 'Precio caja actual

Private mFechanue As String 'fecha nuevo
Private mPrecionu As Currency 'Precio nuevo
Private mPrecion1 As Currency 'Precio caja nuevo

Private mPrecioUPC As Currency

Private mDtopermi As Boolean 'descuento permitido


'ATRIBUTOS DE LA TARIFA (tabla: starif)
Private mMargeCom As Currency
Private mTarifaSobre As Byte
    ' 0.- La tarifa va sobre el PVP
    ' 1.-  "    "  va sobre el UPC




'------------------------------------------------
'Propiedades del modulo CTarifaArt
'------------------------------------------------

'**** Codigo del articulo
Public Property Let CodigoArticulo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get CodigoArticulo() As String
     CodigoArticulo = mCodartic
End Property


'**** Codigo de la tarifa
Public Property Let CodigoTarifa(ByVal vData As Integer)
     mCodLista = vData
End Property

Public Property Get CodigoTarifa() As Integer
     CodigoTarifa = mCodLista
End Property


'**** Precio Actual
Public Property Let PrecioActual(ByVal vData As Currency)
     mPrecioac = vData
End Property

Public Property Get PrecioActual() As Currency
     PrecioActual = mPrecioac
End Property


'**** Precio Caja Actual
Public Property Let PrecioCajaActual(ByVal vData As Currency)
     mPrecioa1 = vData
End Property

Public Property Get PrecioCajaActual() As Currency
     PrecioCajaActual = mPrecioa1
End Property



'**** Fecha cambio nuevo precio
Public Property Let FechaCambio(ByVal vData As String)
     mFechanue = vData
End Property

Public Property Get FechaCambio() As String
     FechaCambio = mFechanue
End Property


'**** Precio Nuevo
Public Property Let PrecioNuevo(ByVal vData As Currency)
     mPrecionu = vData
End Property

Public Property Get PrecioNuevo() As Currency
     PrecioNuevo = mPrecionu
End Property


'**** Precio Caja Nuevo
Public Property Let PrecioCajaNuevo(ByVal vData As Currency)
     mPrecion1 = vData
End Property

Public Property Get PrecioCajaNuevo() As Currency
     PrecioCajaNuevo = mPrecion1
End Property


'**** Descuento permitido (SI/NO)
Public Property Let DescuentoPermitido(ByVal vData As Boolean)
     mDtopermi = vData
End Property

Public Property Get DescuentoPermitido() As Boolean
     DescuentoPermitido = mDtopermi
End Property






'**** Margen Comercial de la tarifa
Public Property Let MargenComercial(ByVal vData As Currency)
     mMargeCom = vData
End Property

Public Property Get MargenComercial() As Currency
     MargenComercial = mMargeCom
End Property

'*****  Tarifa sobre
'           0.- PVP   precio venta public
'           1.- UPC   ultimo precio compra
Public Property Let TarifaSobre(ByVal vData As Byte)
     mTarifaSobre = vData
End Property

Public Property Get TarifaSobre() As Byte
     TarifaSobre = mTarifaSobre
End Property



'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'       METODOS.
'
'       Se cambian a partir del 23 de Abril de 2008
'       Antes el precio en la tarifa se calculaba sobre el PVP aplicando el margen
'       Ahora tb se puede obtener aplicando el margen tarifa sobre el Preccio ult. compra
'
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Public Function LeerDatos(vCodArtic As String, vCodTar As Integer) As Boolean
'Leer los datos de una tarifa de Articulo
'Lee de la BD: Ariges, Tabla: slista
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    'ANTES 23 Abril 2008
    'SQL = "SELECT codartic,slista.codlista,precioac,precioa1,fechanue,precionu,precion1,dtopermi,starif.margecom,opcionINC,preciouc "
    'SQL = SQL & " FROM slista INNER JOIN starif ON slista.codlista=starif.codlista "
    '
    
    SQL = "SELECT slista.codartic,slista.codlista,slista.precioac,slista.precioa1,fechanue,precionu,precion1,dtopermi,"
    SQL = SQL & " starif.margecom , opcionINC, PrecioUC FROM slista INNER JOIN starif ON slista.codlista=starif.codlista"
    SQL = SQL & " INNER JOIN sartic ON slista.codartic=sartic.codartic"
    SQL = SQL & " WHERE slista.codartic=" & DBSet(vCodArtic, "T") & " AND slista.codlista=" & vCodTar
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatos = False
    Else
        CodigoArticulo = CStr(Rs!codArtic)
        CodigoTarifa = Rs!codlista
        
        PrecioActual = Rs!precioac
        PrecioCajaActual = DBLet(Rs!precioa1, "N")
        
        FechaCambio = DBLet(Rs!fechanue, "F")
        PrecioNuevo = DBLet(Rs!precionu, "N")
        PrecioCajaNuevo = DBLet(Rs!precion1, "N")
        
        DescuentoPermitido = Rs!dtopermi
        
        MargenComercial = DBLet(Rs!margecom, "N")
        'Cargo directamente la opcion de tarifa sobre la var
        TarifaSobre = DBLet(Rs!opcionINC, "N")
        mPrecioUPC = DBLet(Rs!precioUC, "N")
        
        
        LeerDatos = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Tarifa Art�culo" & vbCrLf
    SQL = SQL & "N�mero: " & Err.Number & vbCrLf
    SQL = SQL & "Descripci�n: " & Err.Description
    MsgBox SQL, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function



Public Function AplicarMargenComercial(preUC As Currency) As Boolean
'Actualiza los campos temporales de precios con el precio resultado
'de aplicar al precio el margen comercial
Dim Precio As Currency
Dim SQL As String

    On Error GoTo ErrMargen
    
    If mTarifaSobre = 0 Then
        'Sobbre PVP
        Precio = Round2((preUC * Me.MargenComercial) / 100, 4)
        Precio = preUC + Precio
    Else
        'Sobre le ultimo precio de compra
        Precio = Round2((mPrecioUPC * Me.MargenComercial) / 100, 4)
        Precio = mPrecioUPC + Precio
    End If

    SQL = "UPDATE slista SET tmpprecioac=" & DBSet(Precio, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(Me.CodigoArticulo, "T") & " AND codlista=" & Me.CodigoTarifa
    Conn.Execute SQL
    AplicarMargenComercial = True
    Exit Function
    
ErrMargen:
    MuestraError Err.Number, "Aplicar % sobre PVP", Err.Description
End Function



Public Function ActualizarPrecios(newFecha As String, newPrecioAC As Currency, newPrecioA1 As Currency, Optional cadError As String, Optional DentroTRANS As Boolean) As Boolean
Dim SQL As String
Dim numF As String

    If newFecha = "" Then Exit Function
    
    On Error GoTo ErrAct
    If Not DentroTRANS Then Conn.BeginTrans
    
    'Mover los precios actuales al hist�rcio slist1
    '------------------------------------------------
    SQL = "INSERT INTO slist1 (codartic,codlista,numlinea,fechacam,precioac,precioa1) "
    SQL = SQL & " VALUES (" & DBSet(Me.CodigoArticulo, "T") & "," & Me.CodigoTarifa & ","
    'numero de linea
    numF = SugerirCodigoSiguienteStr("slist1", "numlinea", "codartic=" & DBSet(Me.CodigoArticulo, "T") & " AND codlista=" & Me.CodigoTarifa)
    SQL = SQL & numF & "," & DBSet(newFecha, "F") & "," & DBSet(Me.PrecioActual, "N") & "," & DBSet(Me.PrecioCajaActual, "N") & ")"




    Conn.Execute SQL
    

    'Actualizar los precios actuales con valores nuevos
    'y quitar el valor de los precios nuevos y poner a nulos
    '--------------------------------------------------
    SQL = "UPDATE slista SET precioac=" & DBSet(newPrecioAC, "N")
'    SQL = SQL & "," & " precioa1=" & DBSet(newPrecioA1, "N")
    SQL = SQL & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo & ", precion1= " & ValorNulo
    SQL = SQL & " WHERE codartic=" & DBSet(Me.CodigoArticulo, "T") & " AND codlista=" & Me.CodigoTarifa
    Conn.Execute SQL
    
    
    If Not DentroTRANS Then Conn.CommitTrans
    ActualizarPrecios = True
    Exit Function
    
ErrAct:
    ActualizarPrecios = False
    If Not DentroTRANS Then Conn.RollbackTrans
    cadError = cadError & vbCrLf & "Actualizar precios tarifa art�culo."
    cadError = cadError & Err.Number & ": " & Err.Description
'    MuestraError Err.Number, "Actualizar precios tarifa art�culo.", Err.Description
End Function


Private Sub ObtenerMargenComercial()
Dim SQL As String
    On Error GoTo ErrMargen
    
    SQL = DevuelveDesdeBDNew(conAri, "starif", "margecom", "codlista", Me.CodigoTarifa, "N")
    If SQL <> "" Then
        Me.MargenComercial = CCur(SQL)
    End If
    Exit Sub
    
ErrMargen:
    MuestraError Err.Number, "Obtener margen de la tarifa", Err.Description
End Sub




Public Function InsertarPrecios() As Boolean
'al insertar un articulo nuevo insertar la linea para cada una de las tarifas existentes
Dim SQL As String
Dim Precio As Currency

    On Error GoTo ErrIns
    SQL = "INSERT INTO slista (codartic,codlista,precioac,precioa1,fechanue,precionu,precion1,dtopermi,tmpprecioac,tmpprecioa1)"
    SQL = SQL & " VALUES (" & DBSet(Me.CodigoArticulo, "T") & "," & Me.CodigoTarifa & ","
    
    '--- al precio de venta le aplicamos el margen comercial de la tarifa
    'Leer el % de la tarifa
    ObtenerMargenComercial
    'aplicar el % al PVP
    Precio = Round2((Me.PrecioActual * Me.MargenComercial) / 100, 4)
    Precio = Me.PrecioActual + Precio
    
    'precio y precio por caja
    SQL = SQL & DBSet(Precio, "N") & "," & ValorNulo & ","
    'fecha
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1," & ValorNulo & "," & ValorNulo & ")"
    Conn.Execute SQL
    InsertarPrecios = True
    Exit Function
    
ErrIns:
    InsertarPrecios = False
    MuestraError Err.Number, "Insertar lista de precios.", Err.Description
End Function
