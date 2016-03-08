VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'ARTICULOS
'TABLA: prodDepositos




'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mNumDeposito As Integer 'Cod. articulo
Private mCapacidad As Currency
Private mKilos As Currency
Private mLote As String
Private midPartida As Long

Dim SQL As String
'------------------------------------------------
'Propiedades del modulo CArticulo
'------------------------------------------------

'**** NumDeposito
Public Property Let NumDeposito(ByVal vData As Integer)
     mNumDeposito = vData
End Property

Public Property Get NumDeposito() As Integer
     NumDeposito = mNumDeposito
End Property


'**** Capacidad
Public Property Let Capacidad(ByVal vData As Currency)
     mCapacidad = vData
End Property

Public Property Get Capacidad() As Currency
     Capacidad = mCapacidad
End Property




'**** Listros que hay ahora
Public Property Let Kilos(ByVal vData As Currency)
     mKilos = vData
End Property

Public Property Get Kilos() As Currency
     Kilos = mKilos
End Property






'**** Lote mlote
Public Property Let numLote(ByVal vData As String)
     mLote = vData
End Property

Public Property Get numLote() As String
     numLote = mLote
End Property


Public Property Let idPartida(ByVal vData As Long)
     midPartida = vData
End Property

Public Property Get idPartida() As Long
     idPartida = midPartida
End Property




Public Function LeerDatos(NumeroDeposito As Integer, OcutltarMsgbox As Boolean) As Boolean
'Leer los datos de un Articulo dado
'Lee de la BD: Ariges, Tabla: sartic
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT numDeposito,Capacidad,kilos,Numlote ,partida"
    SQL = SQL & " FROM proddepositos "
    SQL = SQL & " WHERE numDeposito=" & CStr(NumeroDeposito)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mNumDeposito = CStr(RS!NumDeposito)
        
        mCapacidad = RS!Capacidad
        mKilos = DBLet(RS!Kilos, "N")
        mLote = DBLet(RS!numLote, "T")
        midPartida = DBLet(RS!Partida, "N") 'Puede ser null o cero
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos deposito" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    If Not OcutltarMsgbox Then MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function


'Requisitos. Se han asignado numdeposito, litros y  Lote
'   0 .- Albaran de compra
'   1 .- Coupage Entrada
'   2 .-  "      salida
'   3 .- Trasiego entrada
'   4 .-    "     salida
'   5 .-  Produccion
'   6 .- Venta directa


Public Function InsertarEnDeposito(Accion As Byte) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = " & DBSet(mKilos, "N")
    'Litros = " & DBSet(mLitros, "N")
    SQL = SQL & ", numlote = " & DBSet(mLote, "T")
    SQL = SQL & ", partida = " & midPartida
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)
    
    InsertarEnDeposito = EjecutaSQL(1, SQL, True)
    
    'DEsde Filtrado llamamos a esta funcion. Pero no queremos que inserte en hco
    If Kilos <> 0 Then InsertarEnHco Accion
    
End Function

'En albaran compra, si puede borrar la linea poneamos a NULL,
'y buscaremos el HCO de acciones

'CodigoInsercioHco: 0 NO    1: Coupage  2: Forzar vaciado
Public Function QuitarAsignacionDeposito_(CodigoInsercioHco As Byte) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = 0, numlote = NULL,Litros=NULL,Partida=NULL"
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)

    QuitarAsignacionDeposito_ = EjecutaSQL(1, SQL, True)
        
        
    If CodigoInsercioHco > 0 Then
        'Es fin cuba de coupage
        If CodigoInsercioHco = 1 Then
            InsertarEnHco 2
        Else
            'Vaciado
            InsertarEnHco 7
        End If
        
    End If
End Function


'Variacion KILos
Public Function VariacionKilosDeposito(Variacion As Currency) As Boolean
    
    Kilos = Kilos + Variacion
    SQL = "UPDATE proddepositos SET Kilos= " & DBSet(Kilos, "N")
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)
    
    VariacionKilosDeposito = EjecutaSQL(1, SQL, True)
        
End Function



Public Function HacerTrasiego(ByRef cDest As cDeposito)
    cDest.Kilos = Me.Kilos
    cDest.idPartida = Me.idPartida
    cDest.numLote = Me.numLote
    
    'Insertamos en hoc la entrada en deposito nuevo
    cDest.InsertarEnDeposito 3
    
    
    QuitarAsignacionDeposito_ 0  'Lo inserto enhco aqui bajo
    InsertarEnHco 4

End Function


'Habra comprobado que los datos del filtrado son correctos
' Es decir , deposito vacios/llenos
' y si ultiza los intermedios tambien
Public Function HacerFiltrado(ByRef cDest As cDeposito, UtilizaIntermedio8 As Boolean, UtilizaIntermedio9 As Boolean, idFiltrado As Long, FechaHoraProceso As Date)
Dim aux As String



    aux = "Filtrado: " & Format(idFiltrado, "000")

    'Salida de este deposito
    InsertarFiltradoHco Me.NumDeposito, 9, aux, FechaHoraProceso

    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio8 Then
        'DEPOSITO 8
        'Entrada
        InsertarFiltradoHco 8, 8, aux, FechaHoraProceso
        
        'Salida
        InsertarFiltradoHco 8, 9, aux, FechaHoraProceso
        
    End If
    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio9 Then
        'DEPOSITO 9
        'Entrada
        InsertarFiltradoHco 9, 8, aux, FechaHoraProceso
        
        'Salida
        InsertarFiltradoHco 9, 9, aux, FechaHoraProceso
    End If



    
    
    'Insertamos en hoc la entrada en deposito nuevo
    'Por culpa de la fecha HORA
    cDest.idPartida = Me.idPartida
    cDest.numLote = Me.numLote
    cDest.Kilos = 0
    cDest.InsertarEnDeposito 0
    InsertarFiltradoHco cDest.NumDeposito, 8, aux, FechaHoraProceso
    cDest.VariacionKilosDeposito Me.Kilos
    
    QuitarAsignacionDeposito_ 0  'Lo inserto enhco aqui bajo
    

End Function






'   0 .- Albaran de compra
'   1 .- Coupage Entrada
'   2 .-  "      salida
'   3 .- Trasiego entrada
'   4 .-    "     salida
'   5 .-  Produccion
'   6 .- Venta directa
'   7 .- Forzar vaciado
'   8 .- FIltrado entrada
'   9 .-   "    salida
Public Function InsertarEnHco(Accion As Byte, Optional LaDescripcion As String)

    'proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion)
    SQL = "INSERT INTO proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion,Descripcion)"
    SQL = SQL & " VALUES (" & mNumDeposito & "," & DBSet(Now, "FH") & "," & midPartida
    SQL = SQL & "," & DBSet(mLote, "T") & "," & Accion & "," & DBSet(LaDescripcion, "T", "S") & ")"
    
    EjecutaSQL conAri, SQL, True
End Function

'Para los filtrados
'Inserta en el hco de acciones, pero con una fecha, y numero deposito distinti
Private Function InsertarFiltradoHco(Depos As Integer, Tipoac As Byte, LaDescripcion As String, ByRef FechaH As Date)

    'proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion)
    SQL = "INSERT INTO proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion,Descripcion)"
    SQL = SQL & " VALUES (" & Depos & "," & DBSet(FechaH, "FH") & "," & midPartida
    SQL = SQL & "," & DBSet(mLote, "T") & "," & Tipoac & "," & DBSet(LaDescripcion, "T", "S") & ")"
    
    EjecutaSQL conAri, SQL, True
    
    'Incrementamos un segundo
    FechaH = DateAdd("s", 1, FechaH)
End Function



