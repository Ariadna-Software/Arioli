VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLineaProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



'variables locales para almacenar los valores de las propiedades
Private mvarCodProduccion As Long 'copia local
Private mvaridLiProd As Long 'copia local
Private mvarLinea As Byte 'copia local
Private mvarLoteTrazabilidad As Long 'copia local
Private mvarcodartic As String 'copia local
Private mvarNomArtic As String 'copia local
Private mvarUdsCaja As Integer
Private mvarLitrosUd As Currency
Private mvarCantidadEstimada As Long


' MU Importante
Private mvarEstado    As Integer  '0.- Pendiente    1.- Iniciado   2.-Modificando  10.-Finalizado

Private mvarFH_Incio As Date
'Private mvarFH_Fin As Date

Private mvarCuantasMP As Byte 'Cunatos subcomponentes tiene
Private vectorComponentes() As cLineaProCompo    'Vector de componentes. No dimensiono entodavia

'Como a la de sin trazabilidad hay que darle tambien salida en el stock
Private MateriaSinTrazabilidad As Collection

Private mvarEtiquetasImpresas As Long


Private mvarFechaCaducidad As String

Private mvarLineaExtraEtiquetas As String
Private mvarLineaExtraEtiqueta2 As String

'Marzo 2013.  para el formato etiqueta de palet
Private mvarTipoImpresionPalet As Byte   '0.- Normal   1.- Olive line

'JUNIO 2014
Private mvarDepo As Integer

Dim SQL As String
Dim vRs As ADODB.Recordset
Dim cL As cLineaProCompo

Dim vTipoMov As CTiposMov  'PARA AUMENTAR el de trazabilidad


Dim TrbConectado As Integer  'Para que no tenga que leerlo 10 veces
Dim vCodAlmac As Integer   'Almacen. Esta en la cabcera de la produccion

Public Property Let NomArtic(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NomArtic = 5
    mvarNomArtic = vData
End Property


Public Property Get NomArtic() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NomArtic
    NomArtic = mvarNomArtic
End Property



Public Property Let codArtic(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codartic = 5
    mvarcodartic = vData
End Property


Public Property Get codArtic() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    codArtic = mvarcodartic
End Property


Public Property Get LoteTrazabilidad() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.LoteTrazabilidad
    LoteTrazabilidad = mvarLoteTrazabilidad
End Property


Public Property Get linea() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.LINEA
    linea = mvarLinea
End Property

Public Property Let linea(ByVal vData As Byte)
    mvarLinea = vData
End Property



Public Property Get idLiProd() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idLiProd
    idLiProd = mvaridLiProd
End Property


Public Property Get CodProduccion() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CodProduccion
    CodProduccion = mvarCodProduccion
End Property


Public Property Get CantidadEstimada() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idLiProd
    CantidadEstimada = mvarCantidadEstimada
End Property

Public Property Get CuantasMP() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idLiProd
    CuantasMP = mvarCuantasMP
End Property


'0.- Pendiente    1.- Iniciado   2.-Modificando 3.-Finalizado
Public Property Get Estado() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idLiProd
    Estado = mvarEstado
End Property


Public Property Get UnidadesCaja() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idLiProd
    UnidadesCaja = mvarUdsCaja
End Property



Public Property Get FH_Incio() As Date
    FH_Incio = mvarFH_Incio
End Property

Public Property Get EtiquetasImpresas() As Long
    EtiquetasImpresas = mvarEtiquetasImpresas
End Property




Public Property Let FechaCaducidad(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codartic = 5
    mvarFechaCaducidad = vData
End Property


Public Property Get FechaCaducidad() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    FechaCaducidad = mvarFechaCaducidad
End Property


Public Property Let LineaExtraEtiquetas(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codartic = 5
    mvarLineaExtraEtiquetas = vData
End Property


Public Property Get LineaExtraEtiquetas() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    LineaExtraEtiquetas = mvarLineaExtraEtiquetas
End Property

Public Property Let LineaExtraEtiqueta2(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codartic = 5
    mvarLineaExtraEtiqueta2 = vData
End Property


Public Property Get LineaExtraEtiqueta2() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    LineaExtraEtiqueta2 = mvarLineaExtraEtiqueta2
End Property



Public Property Let TipoImpresionPalet(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codartic = 5
    mvarTipoImpresionPalet = vData
End Property


Public Property Get TipoImpresionPalet() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    TipoImpresionPalet = mvarTipoImpresionPalet
End Property









'Primero  tendremos que leer de la tabla de lineas de produccion lo que vamos a producir
'Cargar en el vector los componentes y con eso cargado ASIGNAR :
'       - Codigo de trazabilidad
'       - Lote materia prima a TODOS los componentes de trazabilidad
'
'
Public Function AsignarA_LineaProduccion(Deposito As Integer) As Boolean
    'Comprobaremos que no esta nada en la linea
    AsignarA_LineaProduccion = False
    Set vRs = New ADODB.Recordset
    SQL = "Select * from prodtrazlin where cantprodu is null "
    SQL = SQL & " and lineaprod = " & linea
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then
        'MAL ya hay algo en la linea.
        SQL = "Hay produccion en marcha en esa linea:"
        SQL = SQL & "     produccion: " & vRs!Codigo & " / " & vRs!idlin
        SQL = SQL & "     Trazabilidad: " & vRs!lotetraza
        MsgBox SQL, vbExclamation
    Else
        'Insertaremos en las tablas de produccion  . La transaccion estara AQUI
        conn.BeginTrans
        mvarDepo = Deposito
        If PonerArticuloEnCadenaProduccion Then
            conn.CommitTrans
            AsignarA_LineaProduccion = True
        Else
            conn.RollbackTrans
        End If
        
    End If
    vRs.Close
    Set vRs = Nothing
End Function




'***********************************************************
'   para saber si ha cambiado algo
'   TRUE --> Tendremos que volver a leer los datos para volver a pintarlos
'
Public Function LeerDesdeTrazabilidad(CodigoProduccion As Long, idSubLinProd As Long, lineaprod As Byte, lotetraza As Long) As Boolean
Dim J As Integer
Dim Can As Currency

    'Leera de produccion trazabilidad
    'codigo idlin lotetraza lineaprod lotetraza
    Set vRs = New ADODB.Recordset
    SQL = "Select prodtrazlin.*,sartic.codartic,nomartic,cantesti,fhinicio,estado,UniCajas,LitrosUnidad,prodlin.etiquetas,"
    SQL = SQL & " feccaduca,LineaExtraEtiqueta,TipoImpresionPalet,LineaExtraEtiqueta2 from prodtrazlin,prodlin,sartic WHERE"
    SQL = SQL & " prodtrazlin.codigo =prodlin.codigo and prodtrazlin.idlin=prodlin.idlin AND "
    SQL = SQL & " prodtrazlin.lineaprod =" & lineaprod & " and prodtrazlin.lotetraza=" & lotetraza
    SQL = SQL & " AND prodlin.codartic=Sartic.codartic "
    SQL = SQL & " AND prodtrazlin.codigo = " & CodigoProduccion & " AND prodtrazlin.idlin = " & idSubLinProd
    
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    LeerDesdeTrazabilidad = False
    If Not vRs.EOF Then
           mvarcodartic = vRs!codArtic
           mvarCodProduccion = CodigoProduccion
           mvaridLiProd = idSubLinProd
           mvarNomArtic = vRs!NomArtic
           mvarUdsCaja = vRs!Unicajas
           mvarCantidadEstimada = vRs!cantesti
           mvarLinea = lineaprod
           mvarLoteTrazabilidad = lotetraza

           mvarEstado = vRs!Estado  '0.- Pendiente    1.- Iniciado   2.-Modificando  10.-Finalizado
           If IsNull(vRs!fhinicio) Then
                mvarFH_Incio = Now
           Else
                mvarFH_Incio = vRs!fhinicio
           End If
           mvarEtiquetasImpresas = DBLet(vRs!etiquetas, "N")
           
           
                
           mvarFechaCaducidad = DBLet(vRs!feccaduca, "T")
           If mvarFechaCaducidad <> "" Then mvarFechaCaducidad = Format(mvarFechaCaducidad, "dd/mm/yyyy")
           
           LineaExtraEtiquetas = DBLet(vRs!LineaExtraEtiqueta, "T")
           mvarLineaExtraEtiqueta2 = DBLet(vRs!LineaExtraEtiqueta2, "T")
           
           mvarTipoImpresionPalet = DBLet(vRs!TipoImpresionPalet, "N")
           
           LeerDesdeTrazabilidad = True
           
           'Aqui podrai comprobar si hay pero YA  lo hace en el frmde produccion
           
    End If
    vRs.Close
    
    If Not LeerDesdeTrazabilidad Then Exit Function
    
    'Los componentes
    'Como NO esta produciendose entodavia NO los cojo de prodtrazcompo
    'si no que cruzo con codartic
    SQL = "Select prodtrazcompo.*,nomartic,factorconversion,cantidad from prodtrazcompo,sartic,sarti1 WHERE"
    SQL = SQL & " prodtrazcompo.codartic=Sartic.codartic AND "
    SQL = SQL & " sarti1.codarti1=prodtrazcompo.codartic AND sarti1.codartic='" & Me.codArtic & "' AND"
    SQL = SQL & " prodtrazcompo.codigo =" & CodigoProduccion & " AND idlin = " & idSubLinProd
    SQL = SQL & " AND prodtrazcompo.lineaprod =" & lineaprod & " and prodtrazcompo.lotetraza=" & lotetraza
    SQL = SQL & " ORDER BY factorconversion,codartic "
    
    vRs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If vRs.EOF Then
        MsgBox "Error abriendo componentes desde trazabilidad: " & mvarcodartic
    Else
        mvarCuantasMP = 0
        vRs.MoveFirst
        While Not vRs.EOF
            mvarCuantasMP = mvarCuantasMP + 1
            vRs.MoveNext
        Wend
        
        vRs.MoveFirst
        ReDim vectorComponentes(mvarCuantasMP - 1)
        J = 0
        While Not vRs.EOF
            Set vectorComponentes(J) = New cLineaProCompo
            Can = CCur(vRs!FactorConversion)
            Can = Round2(Can * mvarCantidadEstimada, 2)
            If vRs!Cantidad <> 1 Then Can = Round2(Can * vRs!Cantidad, 2)
            vectorComponentes(J).AsignarValoresComponentes vRs!codArtic, vRs!NomArtic, Can, CCur(vRs!FactorConversion), vRs!Cantidad
            vectorComponentes(J).AsignarLote vRs!numLote
            vRs.MoveNext
            J = J + 1
        Wend
        
    End If
    vRs.Close
    Set vRs = Nothing
    
    
    
    
End Function

'***********************************************************
'   para saber si ha cambiado algo
'   TRUE --> Tendremos que volver a leer los datos para volver a pintarlos
'
Public Function LecturaComprobacion() As Boolean

    Set vRs = New ADODB.Recordset
    SQL = "Select * from prodtrazlin where cantprodu is null "
    SQL = SQL & " and lineaprod = " & linea
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then
        'MAL ya hay algo en la linea.
    
    End If
    vRs.Close
End Function



' Lee los datos de la orden de produccion y de sarti1.
' Todavia no esta dentro del proceso de produccion. Con lo caul los datos
'   los cojo de las de las lineas de prod y de la sarti1
'
Public Function LeerDeOrdenProduccion(CodigoProduccion As Long, idSubLinProd As Long) As Boolean
Dim J As Integer
Dim cant As Currency

    'Esto leera a partir de la produccion leera un arituclo de venta para producir
    Set vRs = New ADODB.Recordset
    SQL = "Select prodlin.*,nomartic,UniCajas from prodlin,sartic WHERE prodlin.codartic=Sartic.codartic AND "
    SQL = SQL & " codigo = " & CodigoProduccion & " AND idlin = " & idSubLinProd
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    LeerDeOrdenProduccion = False
    If Not vRs.EOF Then
           mvarcodartic = vRs!codArtic
           mvarCodProduccion = CodigoProduccion
           mvaridLiProd = idSubLinProd
           mvarNomArtic = vRs!NomArtic
           mvarUdsCaja = vRs!Unicajas
           mvarCantidadEstimada = vRs!cantesti
           mvarTipoImpresionPalet = vRs!TipoImpresionPalet
           LeerDeOrdenProduccion = True
    End If
    vRs.Close
    
    If Not LeerDeOrdenProduccion Then Exit Function
    
    'Los componentes
    'Como NO esta produciendose entodavia NO los cojo de prodtrazcompo
    'si no que cruzo con codartic
    SQL = "Select sarti1.*,nomartic,cantidad,factorconversion from sarti1,sartic WHERE sarti1.codarti1=Sartic.codartic AND "
    SQL = SQL & " sarti1.codartic = " & mvarcodartic & " AND trazabilidad=1 ORDER BY factorconversion,codartic"
    vRs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If vRs.EOF Then
        MsgBox "Error abriendo componentes de: " & mvarcodartic
    Else
        mvarCuantasMP = 0
        vRs.MoveFirst
        While Not vRs.EOF
            mvarCuantasMP = mvarCuantasMP + 1
            vRs.MoveNext
        Wend
        
        vRs.MoveFirst
        ReDim vectorComponentes(mvarCuantasMP - 1)
        J = 0
        While Not vRs.EOF
            Set vectorComponentes(J) = New cLineaProCompo
            
            cant = vRs!FactorConversion
            If vRs!Cantidad <> 1 Then cant = cant * vRs!Cantidad
            cant = Round2(cant * mvarCantidadEstimada, 2)
            vectorComponentes(J).AsignarValoresComponentes vRs!codarti1, vRs!NomArtic, cant, CCur(vRs!FactorConversion), vRs!Cantidad
            
            vRs.MoveNext
            J = J + 1
        Wend
        
    End If
    vRs.Close
    Set vRs = Nothing
    
    
    'Reincio resto variables
    mvarLinea = 0
    mvarLoteTrazabilidad = 0
    mvarEstado = 0   '0.- Pendiente    1.- Iniciado   2.-Modificando  10.-Finalizado
    mvarFH_Incio = Now
    mvarFechaCaducidad = ""
    mvarLineaExtraEtiquetas = ""
    mvarLineaExtraEtiqueta2 = ""
End Function


Private Sub CargarArticulosNoTrazabilidad()
    
    Set MateriaSinTrazabilidad = Nothing
    Set MateriaSinTrazabilidad = New Collection
    SQL = "Select sarti1.*,nomartic,cantidad from sarti1,sartic WHERE sarti1.codarti1=Sartic.codartic AND "
    SQL = SQL & " sarti1.codartic = " & mvarcodartic & " AND trazabilidad=0"
    vRs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not vRs.EOF
        'Codartic, nomartic , cantidad
        SQL = vRs!codarti1 & "|" & vRs!NomArtic & "|" & vRs!Cantidad & "|"
        MateriaSinTrazabilidad.Add SQL
        vRs.MoveNext
    Wend
    vRs.Close
    
End Sub



Public Function DevuelveComponenteLinea(V As Integer, ByRef cLinProCompo As cLineaProCompo) As Boolean
    DevuelveComponenteLinea = False
    
    If V > Me.CuantasMP Then Exit Function
    On Error Resume Next
    Set cLinProCompo = vectorComponentes(V - 1)
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Asignado compo. linea"
    Else
        DevuelveComponenteLinea = True
    End If
End Function


Public Function AsignarLoteLinea(V As Integer, LOTEmateriacomponente As String, DesdePistola_1_asignacion As Boolean) As Boolean
'
'  Febrero 2012
'    Puede reservar una linea para producir y luego ir con la pistola y asignar los numeros de lote
'    Por lo tanto la lectura que hara para asignar el lote insertara tb en la BD
'
    
Dim C As Integer

    AsignarLoteLinea = False
    
    If V > Me.CuantasMP Then Exit Function
    On Error Resume Next
    vectorComponentes(V - 1).AsignarLote LOTEmateriacomponente
    
    'Si es insertar en BD es la primera asignacion que se esta haciendo por pistola
    'insertamos en prodtraza componentes
    If DesdePistola_1_asignacion Then
  
            
        ' codigo,idlin,lineaprod,lotetraza,codartic,numlote"
        SQL = "UPDATE prodtrazcompo set numlote = " & DBSet(LOTEmateriacomponente, "T")
        SQL = SQL & " WHERE codigo= " & Me.CodProduccion & " AND idlin= " & Me.idLiProd & " AND lotetraza = " & mvarLoteTrazabilidad
        SQL = SQL & " AND lineaprod = " & Me.linea & " AND codartic = " & DBSet(vectorComponentes(V - 1).codarticCompo, "T")
        
        conn.Execute SQL
        
        SQL = "" 'todas
        For C = 1 To Me.CuantasMP
            If C <> V Then
                If vectorComponentes(C - 1).LoteMateria = "" Then SQL = "NO"
            End If
        Next C
        
        If SQL = "" Then
            'Si ya estan todas las mataux asignadas entonces esta "en produccion"
            SQL = "UPDATE prodlin set estado = 1"
            SQL = SQL & " WHERE codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd
            conn.Execute SQL
        End If
    End If
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Asignado LOTE linea"
    Else
        AsignarLoteLinea = True
    End If
End Function



Private Function PonerArticuloEnCadenaProduccion() As Boolean
Dim R As ADODB.Recordset
Dim I As Byte
Dim NumLoteMateriaAuxAsignadas As Byte
Dim LineaExtra As String

On Error GoTo EPonerArticuloEnCadenaProduccion
    
    PonerArticuloEnCadenaProduccion = False
            
    'Bloqueo EL status de prodlin
    Set R = New ADODB.Recordset
    SQL = "Select estado from prodlin where codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd & " FOR UPDATE"
    R.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If R.EOF Then
        MsgBox "Error bloqueando linea ", vbExclamation
    Else
        If Val(R!Estado) = 0 Then
            '
            '
            Set vTipoMov = New CTiposMov
            vTipoMov.ConseguirContador "TRZ"   'NO pongo control de leer. TIENE QUE EXISTIR
            
            
            If True Then
                Dim DAV As String
                DAV = ""
                Do
                    DAV = InputBox("Introduzca lote trazabilidad", , CStr(vTipoMov.contador))
                    If DAV = "" Then Err.Raise 513, "Debe introducir numero lote"
                    If Not IsNumeric(DAV) Then
                        MsgBox "Debe ser numerico", vbExclamation
                        DAV = ""
                    Else
                        SQL = InputBox("Verificar numero de lote")
                        If SQL <> DAV Then
                            MsgBox "no coinciden", vbExclamation
                            DAV = ""
                        Else
                            'Compruebo si no esta dado de alta en trazabilidad
                            SQL = DevuelveDesdeBD(conAri, "lotetraza", "prodtrazlin", "lotetraza", DAV)
                            If SQL <> "" Then
                                MsgBox "Ya existe la trazibilidad: " & SQL, vbExclamation
                                DAV = ""
                            End If
                        End If
                    End If
                Loop Until DAV <> ""
                'Llegado aqui... funciona
                vTipoMov.contador = Val(DAV)
            
            End If
             mvarLoteTrazabilidad = vTipoMov.contador
            'Insertamo en prodtrazlin
            SQL = "INSERT INTO prodtrazlin(codigo ,idlin ,lineaprod ,lotetraza,DepositoL) VALUES ("
            SQL = SQL & Me.CodProduccion & "," & Me.idLiProd & "," & Me.linea & "," & mvarLoteTrazabilidad & "," & mvarDepo & ")"
            conn.Execute SQL
            
            'insertamos en prodtraza componentes
            SQL = ""
            NumLoteMateriaAuxAsignadas = 0
            For I = 1 To mvarCuantasMP
                ' codigo,idlin,lineaprod,lotetraza,codartic,numlote"
                SQL = SQL & ", (" & Me.CodProduccion & "," & Me.idLiProd & "," & Me.linea & "," & mvarLoteTrazabilidad & ","
                SQL = SQL & DBSet(vectorComponentes(I - 1).codarticCompo, "T") & ","
                SQL = SQL & DBSet(vectorComponentes(I - 1).LoteMateria, "T") & ")"
                If Trim(vectorComponentes(I - 1).LoteMateria) <> "" Then NumLoteMateriaAuxAsignadas = NumLoteMateriaAuxAsignadas + 1
                
            Next
            SQL = Mid(SQL, 2) '1ªcoma fuera
            SQL = "INSERT INTO prodtrazcompo( codigo,idlin,lineaprod,lotetraza,codartic,numlote) VALUES " & SQL
            conn.Execute SQL
            
            'updateo prodlin
            '0.- Pendiente    1.- Iniciado     10.-Finalizado
            SQL = "UPDATE prodlin set estado = "
            If NumLoteMateriaAuxAsignadas = mvarCuantasMP Then
                SQL = SQL & "1"
            Else
                SQL = SQL & "2"
            End If
            SQL = SQL & " , fhinicio = " & DBSet(Now, "FH")
            If IsDate(mvarFechaCaducidad) Then SQL = SQL & " , feccaduca = " & DBSet(mvarFechaCaducidad, "F")
            LineaExtra = mvarLineaExtraEtiquetas
            
            SQL = SQL & " , TipoImpresionPalet = " & mvarTipoImpresionPalet
            
            
            
            'Abril 2013
            'Si es OLIVE LINE, en la etiqueta , si no han indicado nada, pondra
            If mvarTipoImpresionPalet = 1 Then
                If mvarLineaExtraEtiquetas = "" Then
                    'Ejemplo:  L:5778 26/03/2013  EXP:26/09/2014
                    'Marzo2014  EXP: --> BEST BEFORE
                    LineaExtra = "L:" & mvarLoteTrazabilidad & " " & Format(Now, "dd/mm/yyyy")
                    'Buscar B.BEF para ver si actualiz las lineas al cerrar lotes
                    If IsDate(mvarFechaCaducidad) Then LineaExtra = LineaExtra & " B.BEF.:" & Format(mvarFechaCaducidad, "dd/mm/yyyy")
                End If
            End If
            SQL = SQL & " , LineaExtraEtiqueta = " & DBSet(LineaExtra, "T", "S")
            SQL = SQL & " , LineaExtraEtiqueta2 = " & DBSet(mvarLineaExtraEtiqueta2, "T", "S")

            
            SQL = SQL & " WHERE codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd
            conn.Execute SQL
            
            
            vTipoMov.IncrementarContador vTipoMov.TipoMovimiento
            Set vTipoMov = Nothing
            
        End If
        PonerArticuloEnCadenaProduccion = True
    End If
    R.Close
    
EPonerArticuloEnCadenaProduccion:
    If Err.Number <> 0 Then
        If Err.Number = 513 Then
            MsgBox "Proceso cancelado por el usuario", vbExclamation
        Else
            MuestraError Err.Number, "PonerArticuloEnCadenaProduccion" & vbCrLf & Err.Description
        End If
    End If
    Set R = Nothing
    
End Function





Public Function CerrarParaCambioLote(CantidadProducida As Currency, Cajas As Integer, idLINCambioLote As Integer, NuevoLote As String, EsFinalLoteDeposito As Boolean, NumeroDeposito As Integer) As Boolean
Dim b As Boolean
Dim I As Integer
Dim Seguir As Boolean
Dim AntiguaTraza As Long

On Error GoTo ECerrarParaCambioLote

    CerrarParaCambioLote = False
    Set vRs = New ADODB.Recordset
    SQL = "Select estado from prodlin where codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd & " FOR UPDATE"
    vRs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Seguir = False
    If vRs.EOF Then
        MsgBox "Error obteniendo la linea de produccion", vbCritical
    Else
        If Val(vRs!Estado) <> 1 Then
            MsgBox "Estado linea produccion <> 1", vbCritical
        Else
            Seguir = True
        End If
    End If
    vRs.Close
    
    If Seguir Then
            conn.BeginTrans
            
            AntiguaTraza = mvarLoteTrazabilidad  'Para el pase a hco de datos
            mvarDepo = NumeroDeposito
            
            b = ActualizaTrazabilidadLinea(CantidadProducida, Cajas, idLINCambioLote, NuevoLote, EsFinalLoteDeposito)
            
            If b Then
                conn.CommitTrans
                CerrarParaCambioLote = True
                
                InsertarEnHcoProduccion AntiguaTraza
                
            Else
                conn.RollbackTrans
            End If
        End If
    
        If idLINCambioLote > 0 Then
            'Ha cambiado de lote
            'Comprobar que continua paletizando
            If CerrarParaCambioLote Then ComprobarSiguePaletizando
            
        End If
        
        
ECerrarParaCambioLote:
    If Err.Number <> 0 Then MuestraError Err.Number, "CerrarParaCambioLote"
    Set vRs = Nothing
End Function



Public Function CerrarProduccion(CantidadProducida As Currency, Cajas As Integer, EsFinalLoteDeposito As Boolean) As Boolean
Dim b As Boolean
Dim I As Integer
Dim Seguir As Boolean

On Error GoTo ECerrarProduccion

    CerrarProduccion = False
    Set vRs = New ADODB.Recordset
    SQL = "Select estado from prodlin where codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd & " FOR UPDATE"
    vRs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Seguir = False
    If vRs.EOF Then
        MsgBox "Error obteniendo la linea de produccion", vbCritical
    Else
        If Val(vRs!Estado) > 2 Then
            MsgBox "Estado linea produccion >2", vbCritical
        Else
            Seguir = True
        End If
    End If
    vRs.Close
    If Seguir Then
            conn.BeginTrans
            b = ActualizaTrazabilidadLinea(CantidadProducida, Cajas, 0, "", EsFinalLoteDeposito) 'Eso significa cierre produccion
            
            If b Then
                conn.CommitTrans
                CerrarProduccion = True
                
                InsertarEnHcoProduccion mvarLoteTrazabilidad
          
                
                
                'Como ha cerrado la linea de produccion veremos si queda alguna para marcar como producida la cabecera
                'Comprobaremos si todas las sublineas  de la cabecera de produccion estan "producidas"
                'para marcar la cabecera como completada
                SQL = " estado < 10 and codigo "
                SQL = DevuelveDesdeBD(conAri, "count(*)", "prodlin", SQL, Me.CodProduccion)
                If SQL = "" Then SQL = "0"
                If Val(SQL) = 0 Then
                    'YA estan todas producidas
                    SQL = "UPDATE prodcab set producido=1 WHERE codigo= " & Me.CodProduccion
                    EjecutaSQL conAri, SQL
                        
                End If
                
                
                
                
            Else
                conn.RollbackTrans
            End If
        
    End If
    

ECerrarProduccion:
    If Err.Number <> 0 Then MuestraError Err.Number, "CerrarProduccion"
    Set vRs = Nothing
End Function




Private Function ActualizaTrazabilidadLinea(Cantidad As Currency, CajasR As Integer, idLINCambio As Integer, NuevoLotePro As String, FinDepositoLote As Boolean) As Boolean
Dim I As Byte
Dim CanAux As Currency
Dim IdPalet As Long
Dim Cadena As String


Dim cDepo As cDeposito

    On Error GoTo EActualizaTrazabilidadLinea
    ActualizaTrazabilidadLinea = False
     
    
    'Compruebo que existe la linea de trazabilidad
    ' SQL = " prodtrazlin  (codigo ,idlin ,lineaprod ,lotetraza) VALUES ("
    
    SQL = "codigo = " & Me.CodProduccion & " AND idlin = " & Me.idLiProd & " AND lineaprod = " & Me.linea & " AND lotetraza"
    SQL = DevuelveDesdeBD(conAri, "codigo", "prodtrazlin", SQL, CStr(mvarLoteTrazabilidad), "N")
    If SQL = "" Then SQL = "0"
    If Val(SQL) <> Me.CodProduccion Then
        MsgBox "Error inesperado. ActualizaTrazabilidadLinea. Codigos distintos", vbCritical
    Else

        
        'Ok.
        '1.- Updateo la cantidad producida para la linea de produccion
        '2.- Updateo las lineas de componentes con la cantidad utilizada hasta ahora
        
        
        'Junio 2014
        'Actualizo en la tabla de depositos. Tambien veremos si es "final deposito" de
        ' regularizar el lote y vaciar deposito
        
        'Si es CAMBIO DE LOTE, traera un luevo idlin. Si es CERO es que hemos finiquitado el lote
        'para cambio lote seguira por aqui
        '3.- Creo una nueva lineatraza cambiando el codtrazabilidad
        
        'Sea lo que sea, actualizo en smoval y lotes
        
        CargarArticulosNoTrazabilidad   'Los articulos que no llevan TRAZA mueven stock
        
        If Cantidad > 0 Then
            If Not CerrarProduccionCantidades(True, Cantidad, CajasR) Then Exit Function
        End If
        


        SQL = "UPDATE prodtrazlin set cantprodu = " & TransformaComasPuntos(CStr(Cantidad))
        SQL = SQL & " , cajasprod = " & CStr(CajasR)
        SQL = SQL & " WHERE codigo = " & Me.CodProduccion & " AND idlin = " & Me.idLiProd & " AND lineaprod = " & Me.linea & " AND lotetraza = " & mvarLoteTrazabilidad
        conn.Execute SQL


        'Los componentes
        For I = 1 To mvarCuantasMP
            CanAux = Cantidad * vectorComponentes(I - 1).FactorMultiplicador
            If vectorComponentes(I - 1).EsMateriaPrima Then CanAux = Round(CanAux, 2)
            
            
            SQL = "UPDATE prodtrazcompo SET cantutili = " & TransformaComasPuntos(CStr(CanAux))
            ' codigo,idlin,lineaprod,lotetraza,codartic"
            SQL = SQL & " WHERE codigo = " & Me.CodProduccion & " AND idlin = " & Me.idLiProd & " AND lineaprod = " & Me.linea
            SQL = SQL & " AND lotetraza = " & mvarLoteTrazabilidad & " AND codartic = " & DBSet(vectorComponentes(I - 1).codarticCompo, "T")
            conn.Execute SQL
            
            'Para cerrar la produccion
            vectorComponentes(I - 1).CantidadUtilizada = CanAux
            
            
            
            
            'Disminuyo la cantidad en el deposito
            If CanAux <> 0 Then
                If vectorComponentes(I - 1).EsMateriaPrima Then
            
                    Set cDepo = New cDeposito
                    
                    'Vemos en que deposito esta. Para el futuro. Si no fuera materia prima no leeriamos deposito
                    SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", vectorComponentes(I - 1).LoteMateria, "T")
                    If SQL = "" Then Err.Raise 513, , "Error datos deposito"
                
                    cDepo.LeerDatos CInt(SQL), False
                    
                    cDepo.VariacionKilosDeposito -CanAux
                End If
            End If
        Next I





        'Cerramos de verdad actualizando movimientos
        If Cantidad > 0 Then
            If Not CerrarProduccionCantidades(False, Cantidad, CajasR) Then Exit Function
        End If

        If idLINCambio = 0 Then
                Espera 0.2
                
                    
                    
                
                'Si es final de LOTE/DEPOSITO,
                'Hay que ajustar metiendo en smoval/lotes
                ' y si es deposito(Materia axuiliar aun no esta hecho)
                ' vaciamos el deposito
                If FinDepositoLote Then
                    For I = 1 To mvarCuantasMP
                        If vectorComponentes(I - 1).EsMateriaPrima Then
                        
                            RegularizarFinLote_Partida CInt(I)
                            
                        End If
                    Next
                End If
                
                
                
                
                
                
                'CIERRE PRODUCCION. NO HACEMOS NADA
                'Ponemos el estado =10 para saber que esta producido del todo
                SQL = "Codigo = " & Me.CodProduccion & " AND idLin = " & Me.idLiProd & " AND lineaprod "
                SQL = DevuelveDesdeBD(conAri, "sum(cantprodu)", "prodtrazlin", SQL, Me.linea)
                CanAux = CCur(SQL)   'total producido UDS
                
                'Cajas
                SQL = "Codigo = " & Me.CodProduccion & " AND idLin = " & Me.idLiProd & " AND lineaprod "
                SQL = DevuelveDesdeBD(conAri, "sum(cajasprod)", "prodtrazlin", SQL, Me.linea)
                If SQL = 0 Then SQL = "0"
                IdPalet = Val(SQL)   'reutiloz id palet para TOTAL CAJAS PRODUCIDAS
                
                SQL = "UPDATE prodlin set estado = 10 , fhFin = " & DBSet(Now, "FH")
                'Unidades y cajas producidas
                SQL = SQL & " ,cantprodu = " & DBSet(CanAux, "N")
                SQL = SQL & " ,cajasprod = " & CStr(IdPalet)
                SQL = SQL & " WHERE codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd
                conn.Execute SQL
                    
         
                    
                        
            

        Else
        
        
                'Si es final de LOTE/DEPOSITO,
                'Hay que ajustar metiendo en smoval/lotes
                ' y si es deposito(Materia axuiliar aun no esta hecho)
                ' vaciamos el deposito
                If FinDepositoLote Then
                    If vectorComponentes(idLINCambio - 1).EsMateriaPrima Then RegularizarFinLote_Partida idLINCambio
                End If
        
        
                'Veo en que linea se estaba paletizando
                SQL = " ffin is null and lotetraza  "
                SQL = DevuelveDesdeBD(conAri, "idpalet", "prodpaletstraza", SQL, CStr(mvarLoteTrazabilidad))
                IdPalet = 0
                If SQL <> "" Then
                    IdPalet = Val(SQL)
                    'Pongo el palet/traza a finalizado
                    SQL = "UPDATE prodpaletstraza set ffin= " & DBSet(Now, "FH")
                    SQL = SQL & " WHERE lotetraza=" & mvarLoteTrazabilidad & " AND idpalet=" & IdPalet
                    conn.Execute SQL
                    
                    
                End If
                 'Otro numero de trazabilidad
                Set vTipoMov = New CTiposMov
                vTipoMov.ConseguirContador "TRZ"   'NO pongo control de leer. TIENE QUE EXISTIR
                 
                mvarLoteTrazabilidad = vTipoMov.contador
                
                'Insertamo en prodtrazlin
                SQL = "INSERT INTO prodtrazlin(codigo ,idlin ,lineaprod ,lotetraza,depositol) VALUES ("
                SQL = SQL & Me.CodProduccion & "," & Me.idLiProd & "," & Me.linea & "," & mvarLoteTrazabilidad & "," & mvarDepo & ")"
                conn.Execute SQL
        
                'Meto el palet de produccion
                If IdPalet > 0 Then
                    SQL = "INSERT INTO prodpaletstraza (idpalet, lotetraza ,fh) VALUES "
                    SQL = SQL & " (" & IdPalet & "," & mvarLoteTrazabilidad & ",'" & Format(Now, FormatoFechaHora) & "')"
                    conn.Execute SQL
                End If
                
                
                'Actualizo la materia prima de la linea en produccion
                 'insertamos en prodtraza componentes
                SQL = ""
                For I = 1 To mvarCuantasMP
                    
                    ' codigo,idlin,lineaprod,lotetraza,codartic,numlote"
                    SQL = SQL & ", (" & Me.CodProduccion & "," & Me.idLiProd & "," & Me.linea & "," & mvarLoteTrazabilidad & ","
                    SQL = SQL & DBSet(vectorComponentes(I - 1).codarticCompo, "T") & ","
                    If I = idLINCambio Then
                        SQL = SQL & DBSet(NuevoLotePro, "T") & ")"
                    Else
                        SQL = SQL & DBSet(vectorComponentes(I - 1).LoteMateria, "T") & ")"
                    End If
                Next
                SQL = Mid(SQL, 2) '1ªcoma fuera
                SQL = "INSERT INTO prodtrazcompo( codigo,idlin,lineaprod,lotetraza,codartic,numlote) VALUES " & SQL
                conn.Execute SQL
        
                'Updateo en el obejto este
                vectorComponentes(idLINCambio - 1).AsignarLote NuevoLotePro
                
                vTipoMov.IncrementarContador vTipoMov.TipoMovimiento
                
                
                'Abril 2014
                'S pondra
                
                    If mvarLineaExtraEtiquetas <> "" Then
                        If Mid(mvarLineaExtraEtiquetas, 1, 2) = "L:" Then
                            If InStr(1, mvarLineaExtraEtiquetas, "B.BEF.") > 0 Then
                                'Se supone que tenia lo de:
                                 'Ejemplo:  L:5778 26/03/2013  EXP:26/09/2014
                                 'MARZO 2014  EXP--> BEST BEFORE  ---> B.BEF.
                                 Cadena = "L:" & vTipoMov.contador - 1 & " " & Format(Now, "dd/mm/yyyy")
                                 If IsDate(mvarFechaCaducidad) Then Cadena = Cadena & " B.BEF.:" & Format(mvarFechaCaducidad, "dd/mm/yyyy")
                                 mvarLineaExtraEtiquetas = Cadena
                                'UPDATEAMOS
                                SQL = "UPDATE prodlin set LineaExtraEtiqueta = " & DBSet(Cadena, "T", "S")
                                SQL = SQL & " WHERE codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd
                                conn.Execute SQL
                                        
                            
                            End If
                        End If
                    End If
                
                Set vTipoMov = Nothing
        End If
        
        

        
        ActualizaTrazabilidadLinea = True  'Si llega aqui es que ha ido bien
    End If
EActualizaTrazabilidadLinea:
    If Err.Number <> 0 Then MuestraError Err.Number, "ActualizaTrazabilidadLinea", Err.Description
End Function






'*********************************************************************************************
'*********************************************************************************************
'*********************************************************************************************
'
' Cuando cambie de lote o cierre la produccion hara
'       movimientos de entrada/salida en la smoval
'       movimientos de entrada/salida en los LOTES
Private Function CerrarProduccionCantidades(SoloComprobar As Boolean, CantidadProducida_ As Currency, Cajasprod As Integer) As Boolean
Dim vCStock2 As cStock
Dim b As Boolean
Dim cad As String
Dim I As Integer
Dim canti As Currency


    TrbConectado = 1 'Donde llame a este habra que aignar
    vCodAlmac = 1    'ALmacen que habra que leerlo de la cabecera de produccion
    
    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    CerrarProduccionCantidades = False

    Set vCStock2 = New cStock
    
    
    
    b = False

    'Vamos
    For I = 1 To Me.CuantasMP
        b = False
        '                                       no indico cantidad. Ya la leemos del vector
        
        If SoloComprobar Then
            'Teorico
            canti = Round2(CantidadProducida_ * vectorComponentes(I - 1).FactorMultiplicador, 2)
        Else
            canti = vectorComponentes(I - 1).CantidadUtilizada
        End If

        
        If InicializarCStock(vCStock2, "S", CInt(I - 1), canti, "") Then
            
            If vCStock2.MueveStock Then
                If SoloComprobar Then
                    b = vCStock2.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock2.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
                             
        
        If Not b Then I = 9999            'para que se salga
        
        
    Next I
    
    If b Then
        'Los articulos de NO trazabilidad, que stock si que mueven
        For I = 1 To MateriaSinTrazabilidad.Count
            b = False
            
            cad = RecuperaValor(MateriaSinTrazabilidad(I), 3)  'Cantidad de sarti1
            
            canti = CCur(cad)
            canti = Round2(CantidadProducida_ * canti)
        
            cad = RecuperaValor(MateriaSinTrazabilidad(I), 1)
            If InicializarCStock(vCStock2, "S", CInt(I - 1), canti, cad) Then
                
                If vCStock2.MueveStock Then
                    If SoloComprobar Then
                        b = vCStock2.MoverStock(False)
                    Else
                        'Estamos ejecutando la actualizacion
                        '---------------------------------------------
                        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                        'en actualizar stock comprobamos si el articulo tiene control de stock
                        b = vCStock2.ActualizarStock(False)
                    End If
                Else
                    b = True
                End If
            End If
                                 
            
            If Not b Then I = 9999            'para que se salga
            
            
        Next I
    
    End If
    
    
    
    Set vCStock2 = Nothing
      
    If Not b Then
        
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarProduccionLOTES2(SoloComprobar, CantidadProducida_, Cajasprod) Then
            
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If


        Set vCStock2 = New cStock
        b = False
        If InicializarCStock(vCStock2, "E", False, CantidadProducida_, "") Then 'Las lineas son de netrada
            If vCStock2.MueveStock Then
                If SoloComprobar Then
                   ' B = vCStock.MoverStock(False, True)
                   b = True
                Else
                    b = vCStock2.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
        
        
    
    
    If Not b Then
        Set vCStock2 = Nothing
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    CerrarProduccionCantidades = True
    
    
End Function







'Si DesdeMateraNoTrazable tiene valor sera el codartic de la materia axuliar NO trazable (etiquetas,rejillas....)
Private Function InicializarCStock(ByRef vCStock As cStock, TipoM As String, Sublineas As Byte, vcantidadpro As Currency, DesdeMateraNoTrazable As String) As Boolean

    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "TRZ"
    vCStock.Trabajador = TrbConectado
    vCStock.Documento = Format(Me.LoteTrazabilidad, "0000000000")
    vCStock.Fechamov = Me.FH_Incio
    vCStock.codAlmac = vCodAlmac   'LEERLO DESDE cabceraprod
    
    vCStock.Importe = 0
    vCStock.LineaDocu = Me.idLiProd
    vCStock.Cantidad = vcantidadpro
    If TipoM = "E" Then
        'articulo producido
        vCStock.codArtic = Me.codArtic
    Else
        If DesdeMateraNoTrazable <> "" Then
            vCStock.codArtic = DesdeMateraNoTrazable
        Else
            'Materia trazabilidad
            vCStock.codArtic = vectorComponentes(Sublineas).codarticCompo
        End If
    End If
    
    
    
    

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock:" & Err.Description, vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function



Private Function RealizarProduccionLOTES2(SoloComprobar As Boolean, CantidadProducida As Currency, Cajasprod As Integer) As Boolean
Dim ErroresEnPartidas As String
Dim PuedeSeguir As Boolean
Dim cP As cPartidas   'Para los numeros de lote
Dim cL As cLotaje
Dim cad As String
Dim vCan As Currency
Dim botones As Long
Dim N As Integer
Dim K As Byte


    On Error GoTo ERealizarProduccionLOTES

    RealizarProduccionLOTES2 = False
    ErroresEnPartidas = ""

    Set cP = New cPartidas

    If Not SoloComprobar Then

        Set cL = New cLotaje
        cL.DetaMov = "TRZ"
        cL.Documento = Format(Me.LoteTrazabilidad, "0000000000")
        cL.Fechamov = mvarFH_Incio
        cL.HoraMov = CDate(Format(mvarFH_Incio, "dd/mm/yyyy") & " " & Format(Now, "hh:nn:ss"))
        cL.ProvCliTra = TrbConectado
        cL.LineaDocu = Me.CodProduccion
        cL.SubLinea = Me.idLiProd
    End If
        


    PuedeSeguir = True
    For K = 1 To Me.CuantasMP
            cad = ""
            If Not cP.LeerDesdeArticulo(vectorComponentes(K - 1).codarticCompo, vCodAlmac, vectorComponentes(K - 1).LoteMateria) Then
                'MAL ninguna partida asociada al LOTE.
                
                cad = vectorComponentes(K - 1).NomArticCompo & " " & vectorComponentes(K - 1).LoteMateria & "        NO encuetra idPartida asocidado"
                PuedeSeguir = False
                
            Else
                If SoloComprobar Then
                    'Teorico
                    vCan = Round2(CantidadProducida * vectorComponentes(K - 1).FactorMultiplicador, 2)
                Else
                    vCan = vectorComponentes(K - 1).CantidadUtilizada
                End If
                
                If SoloComprobar Then
                    If vCan > cP.Cantidad Then
                        'NECESITA MAS DE LAS QUE VA A PRODUCIR
                        cad = vectorComponentes(K - 1).NomArticCompo
                        If Len(cad) < 30 Then
                            cad = cad & Space(10)
                        ElseIf Len(cad) < 25 Then
                            cad = cad & Space(28)
                        End If
                        cad = cad & "    idPartida: " & cP.idPartida & " Lote: " & vectorComponentes(K - 1).LoteMateria & "        Cantidad INSUFICIENTE (" & vCan & "/" & cP.Cantidad & ")"
                    End If
                Else
                   
                    cP.IncrementarCantidad -1 * vCan
                    
                    InsertarMovientosLotesProduccion cL, cP, vCan, Me.codArtic
                End If
            End If
            
            
            If cad <> "" Then ErroresEnPartidas = ErroresEnPartidas & cad & vbCrLf

    Next K
    
        


   
        
   'ENTRADA DE PRODUCTO
                    'Vemos si ya existe
                    'LoteReal = RL!Numlote & " " & Format(txtFecha(0).Text, "yyyy/mm/dd")
                    If cP.LeerDesdeArticulo(Me.codArtic, vCodAlmac, Format(Me.LoteTrazabilidad, "0000000000")) Then
                        If SoloComprobar Then
                            cad = "     -YA EXISTE el lote de produccion para el articulo de venta:" & vbCrLf
                            cad = cad & Space(20) & Me.NomArtic & "   Lote " & Me.LoteTrazabilidad & "    idPartida: " & cP.idPartida
                            
                            ErroresEnPartidas = cad & vbCrLf & ErroresEnPartidas & vbCrLf
                        Else
                            'Ya existia(por algun motivo)
                            cP.IncrementarCantidad CantidadProducida
                        End If
                        
                    Else
                    
                        'NO EXISTE. NUevo LOTE
                        If Not SoloComprobar Then
                            cP.Cantidad = CantidadProducida
                            cP.codAlmac = vCodAlmac
                            cP.codArtic = Me.codArtic
                            cP.codProve = 0
                            cP.Fecha = FH_Incio
                            'Si cambiamos esto hay que mirar el packing list
                            cP.NumAlbar = "NP" & Format(Me.CodProduccion, "00000") & Format(Me.idLiProd, "00")
                            cP.numLote = Format(Me.LoteTrazabilidad, "0000000000")
                            If Not cP.Insertar Then
                                cad = "Error insertando partidas/lotes: " & cP.codArtic
                                Err.Raise 513, , cad
                            End If
                        
                    End If
                    If Not SoloComprobar Then
                        'En movimientos lote
                        cL.tipoMov = 1
                        cL.Cantidad = cP.Cantidad
                        cL.codAlmac = cP.codAlmac
                        cL.codArtic = cP.codArtic
                        cL.codarti2 = ""
                        cL.numLote = cP.numLote
                        If Not cL.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cP.codArtic
                        
                        
                        '13/12/2010
                        'Vamos a crear la entrada de las cajas.
                        '05 Abril 2011     Las etiquetas NO se crean desde aqui
''''                        Cad = DevuelveDesdeBD(conAri, "max(idcaja)", "prodcajas", "lotetraza", CStr(mvarLoteTrazabilidad))
''''                        If Cad = "" Then Cad = "0"
''''                        botones = Val(Cad)
''''                        Cad = ""
''''                        'prodcajas lotetraza,idcaja,idpalet,fcreacion
''''                        For N = 1 To Cajasprod
''''                            Cad = Cad & ", (" & mvarLoteTrazabilidad & "," & N + botones & ",NULL,NULL)"
''''                        Next N
''''                        Cad = Mid(Cad, 2)
''''                        Cad = "INSERT INTO prodcajas(lotetraza,idcaja,idpalet,fcreacion) VALUES " & Cad
''''                        Conn.Execute Cad
                    End If
                End If
        Set cP = Nothing
   
        If SoloComprobar Then
            If ErroresEnPartidas <> "" Then
                botones = vbExclamation
                If PuedeSeguir Then
                    ErroresEnPartidas = ErroresEnPartidas & vbCrLf & vbCrLf & "¿Continuar?"
                    botones = vbQuestion + vbYesNo
                End If
                        
                If MsgBox(ErroresEnPartidas, botones) = vbNo Then PuedeSeguir = False
            End If
        End If
    
        RealizarProduccionLOTES2 = PuedeSeguir


    
ERealizarProduccionLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description

 
End Function


Private Sub InsertarMovientosLotesProduccion(ByRef cLot As cLotaje, cPar As cPartidas, Cantidad As Currency, ArticuloProduccion As String)

    
    
    cLot.tipoMov = 0  'Salida
    cLot.Cantidad = Abs(Cantidad)
    cLot.codAlmac = cPar.codAlmac
    cLot.codArtic = cPar.codArtic
    cLot.codarti2 = ArticuloProduccion
    cLot.numLote = cPar.numLote

    If Not cLot.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cPar.codArtic
    
End Sub







Public Function LeerLineaDondeEstaPaletizando() As String
    LeerLineaDondeEstaPaletizando = ""
    Set vRs = New ADODB.Recordset
    SQL = "select * from prodpalets where fhfin is null and "
    'Segun la linea
    SQL = SQL & " L" & mvarLinea & " = 1"
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then LeerLineaDondeEstaPaletizando = "L" & vRs!LineaPeletiza & "  - " & Format(vRs!IdPalet, "0000")
    vRs.Close
    Set vRs = Nothing
    
End Function

Public Function CajasLeidasLector() As Integer
    CajasLeidasLector = 0
    Set vRs = New ADODB.Recordset
    SQL = "select count(*) from prodcajas where lotetraza =" & mvarLoteTrazabilidad
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then CajasLeidasLector = DBLet(vRs.Fields(0), "N")
    vRs.Close
    Set vRs = Nothing

End Function




Public Sub EstablecerEtiquetasImpresas(MaximaImpresion As Long)
    mvarEtiquetasImpresas = MaximaImpresion
    SQL = "UPDATE prodlin set etiquetas = " & MaximaImpresion
    SQL = SQL & " WHERE codigo =" & Me.CodProduccion & " AND  idlin  = " & Me.idLiProd
    conn.Execute SQL
End Sub


Private Sub ComprobarSiguePaletizando()
    SQL = "select * from prodpaletstraza  where lotetraza=" & mvarLoteTrazabilidad
    vRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then
        'OK
    
    Else
        'ES EOF, DEBEMOS ASIGNAR LINEA PRODUCCION
        MsgBox "Asigne paletizacion", vbExclamation
        
    End If
    vRs.Close
0
End Sub


Private Sub InsertarEnHcoProduccion(TrazabilidadAnterior As Long)
Dim Aux As String

      '`prodcajasprod`
      'En esta tabla, cada vez que cerramos(cierre o cambio de lote)
      'guardara todos los datos de esa trazabilidad. No tiene mtos
      'ni se crea ni se destruye
      
      Aux = "INSERT INTO `prodcajasprod`(lotetraza,idcaja,idpalet,fcreacion) "
      Aux = Aux & " Select lotetraza,idcaja,idpalet,fcreacion FROM prodcajas WHERE lotetraza = " & TrazabilidadAnterior
      
      If Not EjecutaSQL(conAri, Aux, False) Then
            Aux = "*******  Error insertando en historico  **"
            Aux = Aux & "El programa continuará"
            Aux = Aux & "Avise soporte técnico"
            MsgBox Aux, vbQuestion
      End If
End Sub


'De momento solo entra aqui para materia prima
Private Sub RegularizarFinLote_Partida(IndiceDelArray As Integer)
Dim cPar As cPartidas
Dim cDEP As cDeposito
Dim cLot As cLotaje
'Dim vvCstock2 As cStock
Dim Aux As String
Dim Donde As String
Dim Cantidad As Currency

    On Error GoTo eRegularizarFinLote_Partida


    
    
    Set cPar = New cPartidas
    Set cDEP = New cDeposito
    Set cLot = New cLotaje
    'Set vvCstock = New cStock
    
    Donde = "Leyendo clases"
    Aux = "numlote = " & DBSet(vectorComponentes(IndiceDelArray - 1).LoteMateria, "T") & " AND codartic "
    Aux = DevuelveDesdeBD(conAri, "id", "spartidas", Aux, vectorComponentes(IndiceDelArray - 1).codarticCompo, "T")
    If Aux = "" Then Err.Raise 513, , "No se encuentra la partida"
    cPar.Leer CLng(Aux)
    
    'Vemos en que deposito esta. Para el futuro. Si no fuera materia prima no leeriamos deposito
    Aux = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", vectorComponentes(IndiceDelArray - 1).LoteMateria, "T")
    If Aux = "" Then Err.Raise 513, , "Error datos deposito"
    cDEP.LeerDatos CInt(Aux), False
    
    
    
    
        
    Set cLot = New cLotaje
    
        
   
    
    
    'Un linea mas en smoval
    
    'vvCstock.DetaMov = "DEP"
    '0=Salida, 1=Entrada
    If cPar.Cantidad >= 0 Then
     '   vvCstock.tipoMov = "S"
        cLot.tipoMov = 0
    Else
     '   vvCstock.tipoMov = "E"
        cLot.tipoMov = 1
    End If
    'vvCstock.Cantidad = Abs(cPar.Cantidad)
    'vvCstock.Trabajador = TrbConectado
    ''vCStock.Documento = RecuperaValor(Intercambio, 1)
    'vvCstock.Fechamov = Format(Now, "dd/mm/yyyy")
    'vvCstock.HoraMov = Now
    'vvCstock.codAlmac = vCodAlmac
    'vvCstock.codArtic = cPar.codArtic
    'vvCstock.Importe = 0
    'vvCstock.Documento = "FIN" & Format(cPar.idPartida, "0000000")
    
    cLot.codAlmac = vCodAlmac 'vvCstock.codAlmac
    cLot.codArtic = cPar.codArtic 'vvCstock.codArtic
    cLot.DetaMov = "DEP" 'vvCstock.DetaMov
    cLot.Fechamov = Format(Now, "dd/mm/yyyy") ' vvCstock.Fechamov
    cLot.HoraMov = Now ' vvCstock.HoraMov
    cLot.numLote = cPar.numLote
    
    cLot.Cantidad = Abs(cPar.Cantidad) ' vvCstock.Cantidad
    cLot.LineaDocu = cDEP.NumDeposito
    cLot.Documento = "FIN" & Format(cPar.idPartida, "0000000") 'vvCstock.Documento
    
    cLot.InsertarLote
    'vvCstock.MoverStock False
    cDEP.QuitarAsignacionDeposito2 0, Now, -cDEP.Kilos  'Veremos este codigo
    cPar.AjustarFinPartida
    
    
                        
    
eRegularizarFinLote_Partida:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set cPar = Nothing
    Set cDEP = Nothing
    Set cLot = Nothing
    'Set vvCstock = Nothing
    
End Sub





