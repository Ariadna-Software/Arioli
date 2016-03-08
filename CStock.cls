VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================================================
'MODULO para el control de Stock en Almacenes
'Comprueba si hay stock de un Articulo en un Almacen
'Comprueba si hay control de stock para permitir movimiento o no
'Inserta el movimiento de Entrada/salida de stock en la tabla "smoval"
'Actualiza el valor del stock tras el movimiento en la tabla "salmac"
'=======================================================================


'------------------------------------------------
'Variables del modulo CStock
'------------------------------------------------
Private mvarcodartic As String
Private mvarCodAlmac As Integer

Private mvarFechaMov As String
Private mvarHoraMov As String
Private mvarTipoMov As String '0=Salida, 1=Entrada
Private mvarDetaMov As String 'referencia a codtipom de la tabla stipom
Private mvarCantidad As Single 'Cantidad de Stock a mover
Private mvarImporte As Currency 'Importe de la linea de Movimiento de stock
Private mvarCodOpe As Long   'Codigo Trabajador/Cliente/Proveedor segun el tipo de Movimiento. ENERO 08: LONG
Private mvarDocum As String 'Codigo de la tabla de Movimientos, Traspaso, Albaran,... al que va asociado el movimiento
Private mvarLinea As Integer 'Numero de linea del documento



Private cArticulos As Collection
Private nI As Integer

'variable local
Dim MuestraMenCStock  As Boolean 'Indica si se muestra el mensaje en el control de Stock
                             'Cuando se pasa de Pedido a Albaran si alguna linea no tiene
                             'suficiente stock no se muestra mensaje, al final muestra detalle de articulos
                             'sin stock


'------------------------------------------------
'Propiedades del modulo CStock
'------------------------------------------------

Public Property Let codartic(ByVal vData As String)
     mvarcodartic = vData
End Property

Public Property Get codartic() As String
     codartic = mvarcodartic
End Property


Public Property Let codalmac(ByVal vData As Integer)
     mvarCodAlmac = vData
End Property

Public Property Get codalmac() As Integer
     codalmac = mvarCodAlmac
End Property


Public Property Let Fechamov(ByVal vData As String)
     mvarFechaMov = vData
End Property

Public Property Get Fechamov() As String
     Fechamov = mvarFechaMov
End Property



Public Property Let HoraMov(ByVal vData As String)
     mvarHoraMov = vData
End Property

Public Property Get HoraMov() As String
     HoraMov = mvarHoraMov
End Property


Public Property Let tipoMov(ByVal vData As String)
     mvarTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mvarTipoMov
End Property


Public Property Let DetaMov(ByVal vData As String)
     mvarDetaMov = vData
End Property

Public Property Get DetaMov() As String
     DetaMov = mvarDetaMov
End Property


Public Property Let Cantidad(ByVal vData As Single)
     mvarCantidad = vData
End Property

Public Property Get Cantidad() As Single
     Cantidad = mvarCantidad
End Property


Public Property Let Importe(ByVal vData As Currency)
     mvarImporte = vData
End Property

Public Property Get Importe() As Currency
     Importe = mvarImporte
End Property


Public Property Let Trabajador(ByVal vData As Long)
     mvarCodOpe = vData
End Property

Public Property Get Trabajador() As Long
     Trabajador = mvarCodOpe
End Property


Public Property Let Documento(ByVal vData As String)
     mvarDocum = vData
End Property

Public Property Get Documento() As String
     Documento = mvarDocum
End Property


Public Property Let LineaDocu(ByVal vData As Integer)
     mvarLinea = vData
End Property

Public Property Get LineaDocu() As Integer
     LineaDocu = mvarLinea
End Property



'------------------------------------------------
'Metodos del modulo CStock
'------------------------------------------------


Private Function HayStock2(ByRef ColArt As Collection, vCantidad As Single) As Boolean
'IN: stockOrig: stock existente en almacen Origen
'    stockTras: stock a traspasar del origen a otro almacen
'OUT: True=Hay stock disponible en almacen para realizar el movimiento
'     False=No hay suficiento stock en el almacen
    Dim stockOrig As Single
    Dim stockTras As Single
    Dim B As Boolean

    
    HayStock2 = False
    For nI = 1 To ColArt.Count
        stockOrig = ColArt(nI).cantidadStock
        stockTras = ColArt(nI).CantidadNecesaria * Cantidad
    
        If stockTras > 0 Then
            If stockOrig >= stockTras Then
            'Si cantidad en stock > cantidad a traspasar entonces
                B = True
            Else
                B = False
            End If
        Else
            B = True
        End If
        If Not B Then Exit For
    Next nI
    
    HayStock2 = B
End Function


Private Function ControlStock2(cL As Collection) As Boolean
'Si hay control de Stock no se permite realizar el movimiento
'Sino hay control se pregunta si se permite realizar el movimiento
Dim devuelve As String
Dim B As Boolean
Dim NomArtic As String
        

        devuelve = "Control de Stock  "
        If cL.Count > 1 Then devuelve = devuelve & "(Conjuntos)"
        devuelve = devuelve & ": " & vbCrLf & "----------------------" & vbCrLf & vbCrLf & vbCrLf
        devuelve = devuelve & " No hay suficiente Stock en el almacen  para: "
        For nI = 1 To cL.Count

                
                NomArtic = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", cL(nI).codartic, "T")
                NomArtic = cL(nI).codartic & " - " & NomArtic & " ("
                
                NomArtic = NomArtic & cL(nI).cantidadStock & ")"
                If nI > 1 Then NomArtic = "       L: " & NomArtic
                If nI = 2 Then NomArtic = "----------------------------" & vbCrLf & NomArtic
                devuelve = devuelve & vbCrLf & NomArtic
                
                
                
                
                
        Next nI

        
        If DetaMov = "OFE" Or DetaMov = "PEV" Then 'No hay control de stock para bloquear solo informa
            MsgBox devuelve, vbInformation
            B = False
        Else
            If vParamAplic.ControlStock Then
                'Si hay control Stock no permitir traspaso
                B = True
                If Not MuestraMenCStock Then
                    Select Case DetaMov
                        Case "REG"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Movimiento de Almacen. "
                        Case "TRA"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Traspaso de Almacen. "
                        Case "ALV", "ALR", "ALC"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Albaran. "
                        Case "TRZ"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar la produccion trazabilidad. "
                    End Select
                    MsgBox devuelve, vbExclamation
                End If
            Else
                Select Case DetaMov
                Case "REG"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Movimiento de Almacen? "
                Case "TRA"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Traspaso de Almacen? "
                Case "ALV", "ALR", "ALC"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Albaran? "
                Case "PRO"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar continuar con la producción? "
                Case "TRZ"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar continuar con la producción trazabilidad? "
                End Select
                If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then
                    B = False
                Else
                    B = True
                End If
            End If
        End If
        ControlStock2 = B
End Function


Private Function ObtenerStockNuevo(ByRef ColArticulos As Collection) As Boolean
'OUT: cantidad de Stock que hay en el almacen vColAlm del articulo vCodArtic
'Modificacion  19 Nov.  NO msgbox. Ya que si es conjunto lo mostrare al final
Dim cant As String
Dim devuelve As String

    devuelve = ""
    For nI = 1 To ColArticulos.Count
        Debug.Print ColArticulos.Item(nI).codartic
        cant = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", ColArticulos.Item(nI).codartic, "T", , "codalmac", CStr(codalmac), "N")
        If cant = "" Then 'No hay de ese artículo en Almacen
            cant = vbCrLf & "     .-"
            'Significa que el articulo tiene conjuntos y es del conjunto
            If nI > 1 Then devuelve = devuelve & "(Lin) "
            cant = cant & ColArticulos.Item(nI).codartic.codartic
            devuelve = devuelve & cant
        Else
            'SI que tiene estock
            ColArticulos.Item(nI).cantidadStock = CSng(cant)
            
        End If
    Next
    
    'ANtes solo comprobaba el primero
    'cant = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", codartic, "T", , "codalmac", CStr(codAlmac), "N")
    If devuelve <> "" Then 'No hay de ese artículo(s) en Almacen
        devuelve = "No existe en el Almacen: " & codalmac & vbCrLf
        devuelve = devuelve & "Artículo " & vbCrLf
        MsgBox devuelve, vbExclamation
        ObtenerStockNuevo = False
    Else
        ObtenerStockNuevo = True
    End If
    
End Function


Public Function MueveStock() As Boolean
Dim vArtic As CArticulo
Dim B As Boolean

    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codartic) Then
        If vArtic.EsInstalacion Then
            'Si el articulo es una instalacion no mueve stock
            B = False
        Else
            ' ver si hay control de stock para el articulo
            If vArtic.CtrStock Then
'                    vStock = ObtenerStock
                    B = True
            Else
                B = False
            End If
        End If
   End If
   Set vArtic = Nothing
   MueveStock = B
End Function



Public Function MoverStock(YaTieneLeidoLosDatosObjeto As Boolean, Optional NoMenStock As Boolean) As Boolean
Dim vStock As String
Dim B As Boolean
Dim devuelve As String
Dim vArtic As CArticulo
Dim vD As cLArticulo

    MuestraMenCStock = NoMenStock
       
    MoverStock = False
    
    'Comprobar si hay suficiente stock para realizar el Movimiento de stock si
    'el articulo NO es INSTALACION
    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codartic) Then
        If vArtic.EsInstalacion Then
            B = True
        Else
                        
     
            'meto el primer articulo ppal.
            
            'si hay control de stock para el articulo
            If vArtic.CtrStock Then
                If Not YaTieneLeidoLosDatosObjeto Then
                    Set cArticulos = New Collection
                    Set vD = New cLArticulo
                    vD.codartic = vArtic.Codigo
                    vD.CantidadNecesaria = 1
                    cArticulos.Add vD
                
                    'No se si deberia ir en la clase articulos o en esta
                    If vArtic.EsConjunto And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
                    
                End If
                
                'Aqui leemos del almacen en articulo PPAL
                'vStock = ObtenerStock2
                'If (vStock = "Error") Then Exit Function
                
                'Ademas de comprobar la existencia en BD
                If Not ObtenerStockNuevo(cArticulos) Then
                    If Not YaTieneLeidoLosDatosObjeto Then
                        Set cArticulos = Nothing
                        Set vD = Nothing
                        Exit Function
                    End If
                End If
                
                
                
        
                If HayStock2(cArticulos, Cantidad) Then
                    B = True
                Else 'No hay stock para el movimiento. Ver control de Stock
                    B = Not (ControlStock2(cArticulos))
                 End If
        
            Else
                B = False
            End If

            If Not YaTieneLeidoLosDatosObjeto Then
                Set vD = Nothing
                Set cArticulos = Nothing
            End If
        End If
    End If
    Set vArtic = Nothing
    MoverStock = B
End Function


Public Function ActualizarStock(Optional EsTicket As Boolean) As Boolean
'Actualiza el stock en la tabla smoval y en la tabla salmac
'EsTicket: si estamos generando un ticket no se comprueba el stock auque se quede negativo
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim CadValues As String 'cadena para la SQL de insertar en la tabla salmac
Dim vTipoMov As String
Dim cArt As Collection
Dim vD As cLArticulo
Dim miCantidad As Single

    ActualizarStock = False
    B = False
    On Error GoTo EActualizarStock
    
    
    'Veremos si el articulo es conjunto
    '-------------------------------------------------------------------------------
    Set vD = New cLArticulo
    Set cArticulos = New Collection
    vD.codartic = mvarcodartic
    vD.CantidadNecesaria = 1
    cArticulos.Add vD

    cadMen = DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", mvarcodartic, "T")
    If cadMen = "" Then cadMen = "0"
    If Val(cadMen) = 1 And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
    

    If tipoMov = "S" Then
        vTipoMov = "0" 'Movimiento de Salida de Almacen
    Else
        vTipoMov = "1" 'Movimiento de Entrada de Almacen
    End If
        
    '---- Insertamos en Movimientos Artículos
    ADonde = "Insertando datos en Movimientos de Articulos (tabla: smoval)."
    If Fechamov = "" Then Fechamov = Now
    'Construir la cadena VALUES para insertar en la tabla "smoval"
    For nI = 1 To cArticulos.Count
        
        miCantidad = cArticulos(nI).CantidadNecesaria * Cantidad
        CadValues = "VALUES (" & DBSet(cArticulos(nI).codartic, "T") & ", " & codalmac & ", " & DBSet(Fechamov, "F") & ", '"
        If HoraMov = "" Then HoraMov = Format(Fechamov, "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss")
        CadValues = CadValues & Format(HoraMov, "yyyy-mm-dd hh:mm:ss") & "', " & vTipoMov & " , '" & DetaMov & "', " & DBSet(miCantidad, "N") & ", "
        CadValues = CadValues & DBSet(Importe, "N") & ", " & Trabajador & ", " & ValorNulo & ", " & DBSet(Documento, "T") & ", " & LineaDocu & ")"
        
        B = InsertarMovimArticulos(CadValues)
        
        If Not B Then
            ADonde = ADonde & "Linea: " & nI
            Exit For
        End If
    Next nI
    
    '---- Modificar stock del articulo
    If B Then
        ADonde = "Actualizando Stocks Almacenes (tabla: salmac)."
        If MueveStock Then
            If vTipoMov = "1" Then 'Entrada: no se comprueba si hay stock
'                If ActualizarStockAlmacen() Then b = True
                B = ActualizarStockAlmacen()
            Else
                'If DetaMov = "FTI" Then
                If EsTicket Then
                    'En los tickets se actualiza el stock aunque no haya
                    'no se paraliza la venta
'                    If ActualizarStockAlmacen() Then b = True
                    B = ActualizarStockAlmacen()
                Else
                    'comprobar si hay stock, y si no hay si permito que
                    'se mueva o no.
                    If MoverStock(True) Then
'                        If ActualizarStockAlmacen() Then b = True
                        B = ActualizarStockAlmacen()
                    Else
                        B = False
                    End If
                End If
            End If
        End If
    End If
    
    
    '---- si es articulo de conjunto actualizar stock e insertar movimientos
    'para cada una de las líneas de conjunto
    'if b then b=ActualizarStockLinConjunto
    
    
    ActualizarStock = B
    Exit Function
    
EActualizarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        ActualizarStock = False
        MuestraError Err.Number, cadMen, Err.Description
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String) As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
    SQL = SQL & CadValues
    Conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function


Private Function ActualizarStockAlmacen() As Boolean
Dim SQL As String
Dim ElArticulo As String
Dim miCantidad As Single
On Error Resume Next

    ActualizarStockAlmacen = False
    For nI = 1 To cArticulos.Count
        ElArticulo = DBSet(cArticulos(nI).codartic, "T")
        miCantidad = cArticulos(nI).CantidadNecesaria * Cantidad
        If Cantidad <> 0 Then
            If tipoMov = "E" Then 'Mov. de Entrada
                '==== Aumentar el stock en el Almacen
                SQL = "UPDATE salmac Set canstock = canstock + " & DBSet(miCantidad, "N")
                SQL = SQL & " WHERE codartic =" & ElArticulo & " AND "
                SQL = SQL & " codalmac =" & codalmac
            Else 'Mov. de Salida
                '==== Disminuir Stock en Almacen Origen
                SQL = "UPDATE salmac Set canstock = canstock - " & DBSet(miCantidad, "N")
                SQL = SQL & " WHERE codartic =" & ElArticulo & " AND "
                SQL = SQL & " codalmac =" & codalmac
            End If
            Conn.Execute SQL
        End If
        If Err.Number <> 0 Then
            'Hay error , almacenamos y salimos
            ActualizarStockAlmacen = False
        Else
            ActualizarStockAlmacen = True
        End If
    Next nI
    
End Function


Public Function DevolverStock2() As Boolean
'Cuando se elimina una linea de movimiento de stock:
'Actualiza el stock en la tabla "salmac" aumentando de nuevo la cantidad de la linea
'Eliminar la linea que se introdujo en la "smoval"
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim vD As cLArticulo

    DevolverStock2 = False
    B = False
    
    On Error GoTo EDevolverStock
    
    'Veremos si el articulo es conjunto
    '-------------------------------------------------------------------------------
    Set vD = New cLArticulo
    Set cArticulos = New Collection
    vD.codartic = mvarcodartic
    vD.CantidadNecesaria = 1  'el principal siempre es uno
    cArticulos.Add vD

    cadMen = DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", mvarcodartic, "T")
    If cadMen = "" Then cadMen = "0"
    If Val(cadMen) = 1 And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
        
    
    
    
    'Modificar stock
    ADonde = "Actualizando Stocks Almacenes (tabla: salmac)."
    If MueveStock Then
        If ActualizarStockAlmacen() Then B = True
    Else
        B = True
    End If
    
    If B Then
        'Eliminar en Movimientos Artículos
        ADonde = "Eliminando datos en Movimientos de Articulos (tabla: smoval)."
        If EliminarMovimArticulos Then B = True
    End If
    
    DevolverStock2 = B
    
EDevolverStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        DevolverStock2 = False
    End If
    Set vD = Nothing
    Set cArticulos = Nothing
End Function


Private Function EliminarMovimArticulos() As Boolean
Dim SQL As String
On Error GoTo EElim

    For nI = 1 To cArticulos.Count
        SQL = DBSet(cArticulos(nI).codartic, "T")
        SQL = "DELETE FROM smoval WHERE codartic=" & SQL & " AND codalmac=" & codalmac
        SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
        SQL = SQL & " AND fechamov=" & DBSet(Fechamov, "F")
        'en albaranes de compra añadir el proveedor
        If DetaMov = "ALC" Then
            SQL = SQL & " AND codigope=" & Me.Trabajador 'en compras aqui guardamos el proveedor
        End If
        
        
        Conn.Execute SQL
    
    
    Next nI
    EliminarMovimArticulos = True
    Exit Function
    
EElim:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        EliminarMovimArticulos = False
    
    End If
    Err.Clear
End Function


Public Function ModificarStock2(oldCant As Single) As Boolean
'Cuando se modifica una linea de movimiento de stock:
'Modificar la linea que se introdujo en la "smoval"
'Actualiza el stock en la tabla "salmac" aumentando/disminuyendo la diferencia de cantidad de la linea
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim difCant As Single, difTipoM As String
Dim vD  As cLArticulo

    ModificarStock2 = False
    B = False
    
    On Error GoTo EModificarStock
    
    'Modificar en Movimientos Artículos (tabla: smoval)
    ADonde = "Modificando datos en Movimientos de Articulos (tabla: smoval)."
    If ModificarMovimArticulos Then B = True
    
    If B Then
        
        Set vD = New cLArticulo
        Set cArticulos = New Collection
        vD.codartic = mvarcodartic
        vD.CantidadNecesaria = 1  'el principal siempre es uno
        cArticulos.Add vD
        
        
        'Modificar stock
        ADonde = "Modificando Stocks Almacenes (tabla: salmac)."
        'Obtener si se suma o resta la diferencia de cantidad
        difCant = ObtenerDiferenciaStock(oldCant, difTipoM)
        mvarCantidad = difCant
        mvarTipoMov = difTipoM
        B = ActualizarStockAlmacen()
    End If
    ModificarStock2 = B
    
EModificarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        ModificarStock2 = False
    End If
End Function


Private Function ModificarMovimArticulos() As Boolean
'Modifica la cantidad y el importe en la tabla "smoval"
Dim SQL As String
On Error Resume Next
    
    SQL = "UPDATE smoval SET cantidad=" & DBSet(Cantidad, "N") & ", impormov=" & DBSet(Importe, "N") 'Importe
    SQL = SQL & " WHERE codartic=" & DBSet(codartic, "T") & " AND codalmac=" & codalmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    'NUEVO 23 NOV 2010
    'Si es albaran PROVEEDOR tengo que meter tb el codprove
    If DetaMov = "ALC" Then
        SQL = SQL & " AND codigope = " & Me.Trabajador   'Sera el proveedor
        'Para que no modifuq, por si hubiera, albaranes de años anteriores
        SQL = SQL & " AND fechamov >= '" & Year(Me.Fechamov) & "-01-01'"
    End If
    Conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ModificarMovimArticulos = False
    Else
        ModificarMovimArticulos = True
    End If
End Function


Private Function ObtenerDiferenciaStock(oldCant As Single, difTipoM As String) As Single
'Cuando se modifica la cantidad de una linea hay que reestablecer el stock en la
'tabla "salmac" aumentando o restando la diferencia
'IN -> oldCant: cantidad que habia antes de modificar la cantidad
'OUT -> difTipoMov: Indicara si se devera aumentar o disminuir el stock (Entrada o Salida)
'       RETURN: cantidad a aumentar o disminuir en almacen: salmac
Dim TipoM As Byte
Dim difCant As Single
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EDifStock

    'Obtener el tipo de movimiento en la smoval
    SQL = "SELECT tipomovi FROM smoval "
    SQL = SQL & " WHERE codartic=" & DBSet(codartic, "T") & " AND codalmac=" & codalmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        TipoM = RS.Fields(0).Value
    End If
    RS.Close
    Set RS = Nothing
    
    'En funcion del tipo de Movimiento y la diferencia de stock se sumara o se resta la diferencia de stock
    difCant = oldCant - Cantidad
    If TipoM = 1 Then 'Entrada
        If difCant >= 0 Then 'restar la diferencia
            difTipoM = "S"
        Else 'sumar la diferencia
            difTipoM = "E"
        End If
    ElseIf TipoM = 0 Then 'Salida
        If difCant >= 0 Then
            'sumar la diferencia
            difTipoM = "E"
        Else 'restar la diferencia
            difTipoM = "S"
        End If
    End If
    ObtenerDiferenciaStock = Abs(difCant)
    
EDifStock:
    If Err.Number <> 0 Then
        ObtenerDiferenciaStock = -1 'error
    End If
End Function



'## Laura 22/11/2006
'Private Function ActualizarStockArticConjunto() As Boolean
'Dim cArt As CArticulo
'Dim b As Boolean
'Dim menError As String
'
'    On Error GoTo ErrActConj
'
'    b = False
'    Set cArt = New CArticulo
'    b = cArt.LeerDatos(Me.codArtic) 'articulo q estamos actualizando stock
'    If b Then
'        If cArt.EsConjunto Then
'            'actualizar stock de las lineas de conjunto
'            b = cArt.ActualizarStockLinConjunto(Me, menError)
'        Else
'            b = True 'no es conjunto y salimos con ok
'        End If
'    End If
'    Set cArt = Nothing
'
'    ActualizarStockArticConjunto = b
'    Exit Function
'
'ErrActConj:
'    ActualizarStockArticConjunto = False
'    MuestraError Err.Number, "Actualizar stock artículo de conjunto", Err.Description
'End Function



Private Sub AñadirArticulosConjunto(ByRef vDa As cLArticulo)
Dim RS As ADODB.Recordset
Dim C As String

    
 
    'Si entra aqui es pq no tiene produccion
    '(es decir Morales NO entra, y el AVAB no tiene sctocks de los componentes ya que los
    'componentes se utilizaran para otras cosas, no para dar de baja alta los componentes
     If vParamAplic.EsAVAB Then Exit Sub
     
     
    'No controlo el error. Si da error que pete en la sub que lo llama
    'On Error GoTo EAñadirArticulosConjunto
    
    C = "Select codarti1,cantidad  from sarti1 where codartic = '" & mvarcodartic & "' order by numlinea"
    Set RS = New ADODB.Recordset
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Articulo marcado como conjunto pero no tiene lineas", vbExclamation
    Else
        While Not RS.EOF
            Set vDa = New cLArticulo
            vDa.CantidadNecesaria = RS!Cantidad
            vDa.codartic = RS!codarti1
            cArticulos.Add vDa
            RS.MoveNext
            Set vDa = Nothing
        Wend
    End If
    RS.Close
    Set RS = Nothing

End Sub

Private Sub Class_Terminate()
    Set cArticulos = Nothing
End Sub



