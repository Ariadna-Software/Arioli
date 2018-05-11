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


Private mNoPermiteMolturacion As Boolean 'SI ya se ha envasado NO puede molturar sobre el

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


Public Property Get NoPermiteMolturacion() As Boolean
   NoPermiteMolturacion = mNoPermiteMolturacion
End Property



Public Function LeerDatos(NumeroDeposito As Integer, OcutltarMsgbox As Boolean) As Boolean
'Leer los datos de un Articulo dado
'Lee de la BD: Ariges, Tabla: sartic
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT numDeposito,Capacidad,kilos,Numlote ,partida,NoMolturar"
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
        mLote = DBLet(RS!numLote, "T")
        midPartida = DBLet(RS!partida, "N") 'Puede ser null o cero
        
        mNoPermiteMolturacion = DBLet(RS!nomolturar, "N") = 1
        
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
Public Function InsertarEnDeposito2(Accion As Byte, FechaHora As Date, ObsAux As String) As Boolean
    
    SQL = "UPDATE proddepositos SET Kilos =  " & DBSet(mKilos, "N")
    'Litros = " & DBSet(mLitros, "N")
    SQL = SQL & ", numlote = " & DBSet(mLote, "T")
    SQL = SQL & ", partida = " & midPartida
    
    
    
    If Accion <> 10 Then
        SQL = SQL & ", NoMolturar = 1"
        mNoPermiteMolturacion = False
    End If
    SQL = SQL & ", partida = " & midPartida
    
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)
    
    InsertarEnDeposito2 = EjecutaSQL(1, SQL, True)
    
    If Kilos <> 0 Then InsertarEnHco Accion, FechaHora, ObsAux, mKilos
    
End Function

'En albaran compra, si puede borrar la linea poneamos a NULL,
'y buscaremos el HCO de acciones

'CodigoInsercioHco: 0 NO    1: Coupage (morales-moixent)    2: Forzar vaciado
Public Function QuitarAsignacionDeposito2(CodigoInsercioHco As Byte, FechaHora As Date, Kilos As Currency) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = 0, numlote = NULL,Litros=NULL,Partida=NULL,NoMolturar=0"
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)

    QuitarAsignacionDeposito2 = EjecutaSQL(1, SQL, True)
        
        
    If CodigoInsercioHco > 0 Then
        'Es fin cuba de coupage
        If CodigoInsercioHco = 1 Then
            InsertarEnHco 2, FechaHora, "Fin", Kilos
        Else
            'Vaciado
            InsertarEnHco 7, FechaHora, "Vaciado", Kilos
        End If
        
    End If
End Function




Public Function QuitarAsignacionDepositoVALL(FechaHora As Date, KilosOtroLote As Currency, OtroLote As String) As Boolean
    SQL = "UPDATE proddepositos SET Kilos = 0, numlote = NULL,Litros=NULL,Partida=NULL,NoMolturar=0"
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)

    QuitarAsignacionDepositoVALL = EjecutaSQL(1, SQL, True)
    
    'NO inserta en HCO
End Function



'Variacion KILos
Public Function VariacionKilosDeposito(Variacion As Currency) As Boolean
    
    Kilos = Kilos + Variacion
    SQL = "UPDATE proddepositos SET Kilos= " & DBSet(Kilos, "N")
    SQL = SQL & " WHERE numDeposito=" & CStr(mNumDeposito)
    
    VariacionKilosDeposito = EjecutaSQL(1, SQL, True)
        
End Function



Public Function HacerTrasiego(ByRef cDest As cDeposito, MueveLoteCompleto As Boolean, CuantosKilos As Currency, FechaHora As Date) As Boolean
Dim cP As cPartidas
Dim cL As cLotaje
Dim vC As CTiposMov
Dim J As Integer
Dim Aux As String

    On Error GoTo eHacerTrasiego
    HacerTrasiego = False
    If MueveLoteCompleto Then
    

        If cDest.idPartida = 0 Then
            
            cDest.Kilos = 0
            cDest.idPartida = Me.idPartida
            cDest.numLote = Me.numLote
            cDest.InsertarEnDeposito2 3, FechaHora, "TRASIEGO"
         
        End If
        'Insertamos en hoc la entrada en deposito nuevo
        cDest.VariacionKilosDeposito CuantosKilos
        cDest.InsertarEnHco 3, FechaHora, "Origen " & NumDeposito, CuantosKilos
        
        
        QuitarAsignacionDeposito2 0, FechaHora, 0  'Lo inserto enhco aqui bajo
        InsertarEnHco 4, FechaHora, "Destino " & cDest.NumDeposito, -Me.Kilos

        If Me.NoPermiteMolturacion Then
            cDest.NoPermitirMasMolturacion
        Else
            Aux = "UPDATE proddepositos set NoMolturar =0 WHERE numdeposito = " & cDest.NumDeposito
            EjecutaSQL conAri, Aux
        End If
    Else
        'Va a crear nuevo lotaje, a que no mueve el lote completo
        Set cP = New cPartidas
        Set cL = New cLotaje
         
        If cDest.idPartida = 0 Then
            cDest.Kilos = 0
            cDest.idPartida = Me.idPartida
            cDest.numLote = Me.numLote
            cDest.InsertarEnDeposito2 3, FechaHora, "TRASIEGO"
            Espera 0.25
        End If
        
       
              
        Me.VariacionKilosDeposito -CuantosKilos
        InsertarEnHco 4, FechaHora, "Dest. " & cDest.NumDeposito, -CuantosKilos
        
        
        cDest.VariacionKilosDeposito CuantosKilos
        cDest.InsertarEnHco 3, FechaHora, "Orig " & NumDeposito, CuantosKilos
        
        
        NoPermitirMasMolturacion
        cDest.NoPermitirMasMolturacion
  
        
    End If
    
    
  
        
    HacerTrasiego = True
    
eHacerTrasiego:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set cP = Nothing
    Set cL = Nothing
    Set vC = Nothing
End Function


'Habra comprobado que los datos del filtrado son correctos
' Es decir , deposito vacios/llenos
' y si ultiza los intermedios tambien
Public Function HacerFiltrado(ByRef cDest As cDeposito, UtilizaIntermedio8 As Boolean, UtilizaIntermedio9 As Boolean, idFiltrado As Long, FechaHoraProceso As Date, CuantosKilos As Currency) As Boolean
Dim Aux As String
    
    On Error GoTo eHacerFiltrado
    
    HacerFiltrado = False
    Aux = "Filtrado: " & Format(idFiltrado, "000")
    
    'Esto sera en ACEITES MORALES
    If CuantosKilos = 0 Then CuantosKilos = Me.Kilos

    'Salida de este deposito
    InsertarFiltradoHco3 Me.NumDeposito, 9, Aux, FechaHoraProceso, -CuantosKilos, Kilos - CuantosKilos
  

    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio8 Then
        'DEPOSITO 8
        'Entrada
        InsertarFiltradoHco3 8, 8, Aux, FechaHoraProceso, CuantosKilos, 0
        
        
        'Salida
        InsertarFiltradoHco3 8, 9, Aux, FechaHoraProceso, CuantosKilos, 0
        
        
    End If
    'Si se utilizan intermedios, insertaremos en HCO de depositos, uno de entrada y otro de salida 8separados un seg
    If UtilizaIntermedio9 Then
        'DEPOSITO 9
        'Entrada
        InsertarFiltradoHco3 9, 8, Aux, FechaHoraProceso, CuantosKilos, 0
       
        'Salida
        InsertarFiltradoHco3 9, 9, Aux, FechaHoraProceso, CuantosKilos, 0
       
    End If



    
    
    'Insertamos en hoc la entrada en deposito nuevo
    'Por culpa de la fecha HORA
    If cDest.idPartida = 0 Then
        'Asignamos la nueva partida
        cDest.idPartida = Me.idPartida
        cDest.numLote = Me.numLote
        cDest.Kilos = 0
        
    
        cDest.InsertarEnDeposito2 0, FechaHoraProceso, Aux
    Else
        
        
    End If
    
    
    If CuantosKilos >= Kilos Then
        'Ha movido todo el deposito
        If cDest.NumDeposito <> Me.NumDeposito Then
            QuitarAsignacionDeposito2 0, FechaHoraProceso, -Me.Kilos 'Lo inserto enhco aqui bajo
        Else
            'Es le mismo deposito. Lo pongo a cero para que bajo sume la cantidad
            cDest.Kilos = 0
        End If
    Else
        'Ha movido solo parte de los datos
        VariacionKilosDeposito -CuantosKilos
        
    End If
    InsertarFiltradoHco3 cDest.NumDeposito, 8, Aux, FechaHoraProceso, CuantosKilos, cDest.Kilos + CuantosKilos
    cDest.VariacionKilosDeposito CuantosKilos
    
    
    
    'YA no puede molturar sobre ninguno de los dos depositos
    
    NoPermitirMasMolturacion
    cDest.NoPermitirMasMolturacion
    HacerFiltrado = True
    
    Exit Function
eHacerFiltrado:
    
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
' 11.-  Regularizacion
Public Function InsertarEnHco(Accion As Byte, ByVal Fecha As Date, LaDescripcion As String, Cantidad As Currency)


    If UltimaEntradaEnHcoDepositos = Fecha Then
        Fecha = DateAdd("s", 1, Fecha)
    End If
        
    
    'proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion)
    SQL = "INSERT INTO proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion,Descripcion,CantidadMov,stockDep)"
    SQL = SQL & " VALUES (" & mNumDeposito & "," & DBSet(Fecha, "FH") & "," & midPartida
    SQL = SQL & "," & DBSet(mLote, "T") & "," & Accion & "," & DBSet(LaDescripcion, "T", "S")
    SQL = SQL & "," & DBSet(Cantidad, "N") & "," & DBSet(Kilos, "N") & ")"
    
    EjecutaSQL conAri, SQL, True
    UltimaEntradaEnHcoDepositos = Fecha
End Function

'Para los filtrados
'Inserta en el hco de acciones, pero con una fecha, y numero deposito distinti
Private Function InsertarFiltradoHco3(Depos As Integer, Tipoac As Byte, LaDescripcion As String, ByRef FechaH As Date, Kilos2 As Currency, KilosKedan As Currency)

    'proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion)
    SQL = "INSERT INTO proddepositoshco(numDeposito,horamovi,Partida,Numlote,tipoaccion,Descripcion,CantidadMov,stockDep)"
    SQL = SQL & " VALUES (" & Depos & "," & DBSet(FechaH, "FH") & "," & midPartida
    SQL = SQL & "," & DBSet(mLote, "T") & "," & Tipoac & "," & DBSet(LaDescripcion, "T", "S")
    SQL = SQL & "," & DBSet(Kilos2, "N") & "," & DBSet(KilosKedan, "N") & ")"
    
    EjecutaSQL conAri, SQL, True
    
    'Incrementamos un segundo
    FechaH = DateAdd("s", 1, FechaH)
End Function


Public Sub NoPermitirMasMolturacion()
    SQL = "UPDATE proddepositos set nomolturar=1 WHERE numdeposito=" & Me.NumDeposito
    EjecutaSQL conAri, SQL, True
    
End Sub


Public Function AvisarMovimientoHcoPosterior(Fecha As Date) As Boolean
    AvisarMovimientoHcoPosterior = False
    SQL = DevuelveDesdeBD(conAri, "max(horamovi)", "proddepositoshco", "numdeposito", CStr(Me.NumDeposito))
    If SQL = "" Then Exit Function
    
    If CDate(SQL) > Fecha Then
        
        
        'Hay un movimiento posterior a la fecha
        AvisarMovimientoHcoPosterior = True
     
    End If
       
End Function

