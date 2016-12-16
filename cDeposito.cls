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
Public Property Let NUmlote(ByVal vData As String)
     mLote = vData
End Property

Public Property Get NUmlote() As String
     NUmlote = mLote
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
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mNumDeposito = CStr(RS!NumDeposito)
        
        mCapacidad = RS!Capacidad
        mKilos = DBLet(RS!Kilos, "N")
        mLote = DBLet(RS!NUmlote, "T")
        midPartida = DBLet(RS!partida, "N") 'Puede ser null o cero
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


Public Function InsertarEnDeposito(Accion As Byte, FechaHora As Date) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = " & DBSet(mKilos, "N")
    'Litros = " & DBSet(mLitros, "N")
    SQL = SQL & ", numlote = " & DBSet(mLote, "T")
    SQL = SQL & ", partida = " & midPartida
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)
    
    InsertarEnDeposito = EjecutaSQL(1, SQL, True)
    
    'DEsde Filtrado llamamos a esta funcion. Pero no queremos que inserte en hco
    If Kilos <> 0 Then InsertarEnHco Accion, FechaHora
    
End Function

'En albaran compra, si puede borrar la linea poneamos a NULL,
'y buscaremos el HCO de acciones

'CodigoInsercioHco: 0 NO    1: Coupage  2: Forzar vaciado
Public Function QuitarAsignacionDeposito_(CodigoInsercioHco As Byte, FechaHora As Date) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = 0, numlote = NULL,Litros=NULL,Partida=NULL"
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)

    QuitarAsignacionDeposito_ = EjecutaSQL(1, SQL, True)
        
        
    If CodigoInsercioHco > 0 Then
        'Es fin cuba de coupage
        If CodigoInsercioHco = 1 Then
            InsertarEnHco 2, FechaHora
        Else
            'Vaciado
            InsertarEnHco 7, FechaHora
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



Public Function HacerTrasiego(ByRef cDest As cDeposito, MueveLoteCompleto As Boolean, CuantosKilos As Currency, FechaHora As Date) As Boolean
Dim Cp As cPartidas
Dim cL As cLotaje
Dim vC As CTiposMov
Dim J As Integer

    On Error GoTo eHacerTrasiego
    HacerTrasiego = False
    If MueveLoteCompleto Then
    
        'El que habia. Mueve el lote completo. No creaa nuevo lotaje
        cDest.Kilos = Me.Kilos
        cDest.idPartida = Me.idPartida
        cDest.NUmlote = Me.NUmlote
        
        'Insertamos en hoc la entrada en deposito nuevo
        cDest.InsertarEnDeposito 3, FechaHora
        
        
        QuitarAsignacionDeposito_ 0, FechaHora 'Lo inserto enhco aqui bajo
        InsertarEnHco 4, FechaHora

    
    Else
        'Va a crear nuevo lotaje, a que no mueve el lote completo
        Set Cp = New cPartidas
        Set cL = New cLotaje
         
        'NO VA A MOVER STOCK, ya que el articulo es el mismo. Lo unico es que lo pasa de un deposito a otro
        'Creando un numero de lote
        
        Cp.Cantidad = CuantosKilos
        Cp.codAlmac = 1
        Cp.codartic = DevuelveDesdeBD(conAri, "codartic", "spartidas", "id", Me.idPartida)
        Cp.codProve = 0
        Cp.Fecha = CDate(FechaHora)
        Set vC = New CTiposMov
        If Not vC.Leer("TRS") Then Err.Raise 513, "Conseguir contador"
        Cp.NumAlbar = "TRS" & Format(vC.contador + 1, "0000000")
                
        J = InStrRev(Me.NUmlote, "T")
        If J < 5 Then J = 0
        If J > 0 Then
            
            SQL = Mid(Me.NUmlote, J + 1)
            If Not IsNumeric(SQL) Then Err.Raise "ERROR Obteniendo nuevo lote (Trasiego)"
            SQL = Mid(Me.NUmlote, 1, J) & Val(SQL) + 1
            
        Else
            SQL = Me.NUmlote & "T1"
        End If
         
        Cp.NUmlote = SQL
        If Not Cp.Insertar Then Err.Raise 513, "Error insertando partidas/lotes: " & Cp.codartic
                                    
        cDest.Kilos = CuantosKilos
        cDest.idPartida = Cp.idPartida
        cDest.NUmlote = Cp.NUmlote
                                    
                                    
        
        cL.DetaMov = vC.TipoMovimiento
        cL.Documento = Cp.NumAlbar
        cL.Fechamov = Cp.Fecha
        cL.HoraMov = FechaHora
        cL.ProvCliTra = "1"
        cL.LineaDocu = 0
        cL.SubLinea = 0
        cL.codartic = Cp.codartic
        cL.codAlmac = Cp.codAlmac
        cL.tipoMov = 1
        cL.Cantidad = CuantosKilos
        cL.NUmlote = Cp.NUmlote
        If Not cL.InsertarLote Then Err.Raise 513, "Insertar lote1"
        
        'Damos de baja en el lote anterior
        cL.tipoMov = 0
        cL.NUmlote = Me.NUmlote
        Espera 0.5
        If Not cL.InsertarLote Then Err.Raise 513, "Insertar lote"
        'Damos de baja en la partida
        Set Cp = Nothing
        Set Cp = New cPartidas
        If Not Cp.Leer(Me.idPartida) Then Err.Raise 513, "Error insertando partidas/lotes: " & idPartida
       
        Cp.IncrementarCantidad -CuantosKilos
        
        'Ahora los depositos
        
        
        
        
        
        
        
        'Ahora daremos de baja
        
        'Insertamos en hoc la entrada en deposito nuevo
        cDest.InsertarEnDeposito 3, FechaHora
        
        
        Me.VariacionKilosDeposito -CuantosKilos
        InsertarEnHco 4, FechaHora
        
        
        vC.IncrementarContador vC.TipoMovimiento
        
        
    End If
    HacerTrasiego = True
    
eHacerTrasiego:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Cp = Nothing
    Set cL = Nothing
    Set vC = Nothing
End Function


'Habra comprobado que los datos del filtrado son correctos
' Es decir , deposito vacios/llenos
' y si ultiza los intermedios tambien
Public Function HacerFiltrado(ByRef cDest As cDeposito, UtilizaIntermedio8 As Boolean, UtilizaIntermedio9 As Boolean, idFiltrado As Long, FechaHoraProceso As Date)
Dim Aux As String



    Aux = "Filtrado: " & Format(idFiltrado, "000")

    'Salida de este deposito
    InsertarFiltradoHco Me.NumDeposito, 9, Aux, FechaHoraProceso
  

    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio8 Then
        'DEPOSITO 8
        'Entrada
        InsertarFiltradoHco 8, 8, Aux, FechaHoraProceso
        
        
        'Salida
        InsertarFiltradoHco 8, 9, Aux, FechaHoraProceso
        
        
    End If
    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio9 Then
        'DEPOSITO 9
        'Entrada
        InsertarFiltradoHco 9, 8, Aux, FechaHoraProceso
       
        'Salida
        InsertarFiltradoHco 9, 9, Aux, FechaHoraProceso
       
    End If



    
    
    'Insertamos en hoc la entrada en deposito nuevo
    'Por culpa de la fecha HORA
    cDest.idPartida = Me.idPartida
    cDest.NUmlote = Me.NUmlote
    cDest.Kilos = 0
    cDest.InsertarEnDeposito 0, FechaHoraProceso
    InsertarFiltradoHco cDest.NumDeposito, 8, Aux, FechaHoraProceso
    cDest.VariacionKilosDeposito Me.Kilos
    
    QuitarAsignacionDeposito_ 0, FechaHoraProceso 'Lo inserto enhco aqui bajo
    

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
'  10 .- Moltruaracion almazar
Public Function InsertarEnHco(Accion As Byte, Fecha As Date, Optional LaDescripcion As String)

    'proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion)
    SQL = "INSERT INTO proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion,Descripcion)"
    SQL = SQL & " VALUES (" & mNumDeposito & "," & DBSet(Fecha, "FH") & "," & midPartida
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



