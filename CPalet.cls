VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPalet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const NumeroLineasProd = 7 '0+1

'variables locales para almacenar los valores de las propiedades
Private mvarID As Long 'copia local
Private mvarLineaPeletizacion As Byte  'de 1 a 3
Private mvarFechaInicio As Date 'copia local
Private mvarFechaFin As Date 'copia local
                                    
Private mvarLineas(NumeroLineasProd) As Boolean
                                    
Private mvaridTrazabilidad As String  'Cuando creamos un palet añadimos el idtraza que estoy producciendo en esa linea

Private mvarTipoImpresion As Byte  '0.- Normal   1,. Oliveline
                                    
Dim SQL As String
Dim J As Byte
Dim vRs As ADODB.Recordset
                                    
                                    





Public Property Get LineasProd(Index As Integer) As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CadenaLineas
    LineasProd = mvarLineas(Index)
End Property

Public Property Let LineasProd(Index As Integer, ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaFin = 5
    mvarLineas(Index) = vData
End Property


Public Property Let FechaFin(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaFin = 5
    mvarFechaFin = vData
End Property


Public Property Get FechaFin() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaFin
    FechaFin = mvarFechaFin
End Property



Public Property Let FechaInicio(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaInicio = 5
    mvarFechaInicio = vData
End Property


Public Property Get FechaInicio() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaInicio
    FechaInicio = mvarFechaInicio
End Property



Public Property Let ID(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



Public Property Let LineaPeletizacion(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ID = 5
    mvarLineaPeletizacion = vData
End Property


Public Property Get LineaPeletizacion() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID
    LineaPeletizacion = mvarLineaPeletizacion
End Property


Public Property Get TrazabilidadPaletizando() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID
    TrazabilidadPaletizando = mvaridTrazabilidad
End Property


Public Property Let TipoImpresion(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ID = 5
    mvarTipoImpresion = vData
End Property


Public Property Get TipoImpresion() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID
    TipoImpresion = mvarTipoImpresion
End Property





Public Function Leer(vID As Long) As Boolean
    
    On Error GoTo ELeer
    Leer = False
    SQL = "Select * from prodpalets WHERE idpalet = " & CStr(vID)
    Set vRs = New ADODB.Recordset
    vRs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not vRs.EOF Then
        mvarID = vID
        mvarFechaInicio = vRs!fhinicio
        mvarLineaPeletizacion = vRs!LineaPeletiza
        If Not IsNull(vRs!fhFin) Then
            mvarFechaFin = vRs!fhFin
        End If
        mvarLineas(0) = vRs!L0
        mvarLineas(1) = vRs!L1
        mvarLineas(2) = vRs!L2
        mvarLineas(3) = vRs!L3
        mvarLineas(4) = vRs!L4
        mvarLineas(5) = vRs!L5
        mvarLineas(6) = vRs!L6
        mvarLineas(7) = vRs!L7
        
        mvarTipoImpresion = DBLet(vRs!TipoImpresion, "N")
        
        Leer = True
    End If
    vRs.Close
    If Leer Then
        SQL = "select lotetraza from prodpaletstraza where  prodpaletstraza.idpalet =" & mvarID
        SQL = SQL & " AND ffin is null"
        vRs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        mvaridTrazabilidad = ""
        While Not vRs.EOF
            mvaridTrazabilidad = mvaridTrazabilidad & vRs!lotetraza & "|"
            vRs.MoveNext
        Wend
        vRs.Close
    End If

    
ELeer:
    If Err.Number <> 0 Then MuestraError Err.Number, , SQL
    Set vRs = Nothing
End Function


Public Function TodasLasTrazabilidades() As Boolean


    'Para cerrar palets. Puede cerrar con cualquiera de las cajas de palet
    'aunque YA no este paletizandose
    
        SQL = "select lotetraza from prodpaletstraza where  prodpaletstraza.idpalet =" & mvarID
        SQL = SQL & " AND not ffin is null"
        Set vRs = New ADODB.Recordset
        vRs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       
        While Not vRs.EOF
            mvaridTrazabilidad = mvaridTrazabilidad & vRs!lotetraza & "|"
            vRs.MoveNext
        Wend
        vRs.Close
        Set vRs = Nothing
End Function



Public Function CrearPalet() As Boolean
    
    On Error GoTo ELeer
    CrearPalet = False
    Set vRs = New ADODB.Recordset
    SQL = "Select max(idpalet) from prodpalets FOR UPDATE"
    
    vRs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = DBLet(vRs.Fields(0), "N")
    If SQL = "" Then SQL = "0"
    SQL = Val(SQL) + 1
    Me.ID = Val(SQL)
    SQL = "INSERT INTO prodpalets(idpalet,LineaPeletiza,fhinicio,L0,L1,L2,L3,L4,L5,L6,L7,TipoImpresion) VALUES (" & SQL & ","
    SQL = SQL & Me.LineaPeletizacion & "," & DBSet(Me.FechaInicio, "FH")
    For J = 0 To NumeroLineasProd
        SQL = SQL & "," & DBSet(mvarLineas(J), "B")
    Next J
    SQL = SQL & "," & mvarTipoImpresion & ")"
    
    If EjecutaSQL(conAri, SQL, True) Then CrearPalet = True
    vRs.Close
    
    
    
    'Añadimos en las lineas
    'prodpaletstraza idpalet lotetraza fh
    SQL = ""
    While mvaridTrazabilidad <> ""
        J = InStr(1, mvaridTrazabilidad, "|")
        If J = 0 Then
            mvaridTrazabilidad = ""
        Else
            SQL = SQL & ", (" & Me.ID & "," & Mid(mvaridTrazabilidad, 1, J - 1) & ",'" & Format(Now, FormatoFechaHora) & "')"
            mvaridTrazabilidad = Mid(mvaridTrazabilidad, J + 1)
        End If
    Wend
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'quito la primera coma
        SQL = "INSERT INTO prodpaletstraza (idpalet, lotetraza ,fh) VALUES " & SQL
        Conn.Execute SQL
    End If
    
ELeer:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vRs = Nothing
End Function


Public Function CerrarPalet(TotalCajas As Integer) As Boolean
      On Error Resume Next
      
    'El palet de trazanilidad lo pongo a fin
    SQL = "UPDATE prodpaletstraza set fFin = " & DBSet(Now, "FH") & " WHERE idpalet= " & mvarID
    Conn.Execute SQL
    If Err.Number <> 0 Then Err.Clear
    
    SQL = "UPDATE prodpalets set fhFin = " & DBSet(Now, "FH") & " , CajasProd=" & TotalCajas
    SQL = SQL & " WHERE idpalet= " & mvarID
    Conn.Execute SQL
    
   
   
    
    If Err.Number <> 0 Then
        MuestraError Err.Number
        CerrarPalet = False
    Else
        CerrarPalet = True
    End If
End Function



'Reabrir palet
Public Function ReAbrirPalet(NuevasTrazabilidades As String) As Boolean
      On Error Resume Next
      
      
    
    
    'Añadimos en las trzabilidad
    'prodpaletstraza idpalet lotetraza fh
    SQL = ""
    While NuevasTrazabilidades <> ""
        J = InStr(1, NuevasTrazabilidades, "|")
        If J = 0 Then
            NuevasTrazabilidades = ""
        Else
            SQL = SQL & ", (" & Me.ID & "," & Mid(NuevasTrazabilidades, 1, J - 1) & ",'" & Format(Now, FormatoFechaHora) & "')"
            NuevasTrazabilidades = Mid(NuevasTrazabilidades, J + 1)
        End If
    Wend
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'quito la primera coma
        SQL = "INSERT INTO prodpaletstraza (idpalet, lotetraza ,fh) VALUES " & SQL
        Conn.Execute SQL
    End If
    
    
  
    
    
   
    SQL = "UPDATE prodpalets set fhFin = NULL , CajasProd=0, LineaPeletiza = " & Me.LineaPeletizacion
    
    'Lineas
    For J = 0 To NumeroLineasProd
        SQL = SQL & ", L" & J & " = " & DBSet(mvarLineas(J), "B")
    Next J
    
    SQL = SQL & " WHERE idpalet= " & mvarID
    Conn.Execute SQL
    
   
   
    
    If Err.Number <> 0 Then
        MuestraError Err.Number
        ReAbrirPalet = False
    Else
        ReAbrirPalet = True
    End If
End Function


'Mostrar cajas
'Dado un list carga las cajas
Public Sub CargaCajasPaletList(ByRef Total As Integer, ByRef Lis As ListBox)
    On Error GoTo ECargaCajasPaletList
    Set vRs = New ADODB.Recordset
    SQL = "select * from prodcajas where idpalet = " & mvarID
    'Abril 2012. Que odene desdendente
    SQL = SQL & " ORDER BY lotetraza desc,idcaja desc"
    
    vRs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Lis.Clear
    Total = 0
    While Not vRs.EOF
        SQL = Format(vRs!lotetraza, "00000000") & Format(vRs!idcaja, "00000")
        Lis.AddItem SQL
        Total = Total + 1
        vRs.MoveNext
    Wend
    vRs.Close

ECargaCajasPaletList:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set vRs = Nothing
End Sub


'Devolvera una collection que tendra:
'codartic|nomartic|lotetraza|cajas
'MultiplesPaletsImpresion: No borrara la tmp
Public Sub CargaDatosPalet(ByRef Col As Collection, PreparaImpresion As Boolean, ByRef TotalCajas As Integer, MultiplesPaletsImpresion As Boolean)
Dim RN As ADODB.Recordset
Dim cad As String
Dim F As Date

    On Error GoTo ECargaDatosPalet

    Set vRs = New ADODB.Recordset
    TotalCajas = 0
    Set Col = New Collection
    SQL = "Select lotetraza,count(*)  from prodcajas where idpalet = " & mvarID & " GROUP BY 1"
    vRs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    SQL = ""
    While Not vRs.EOF
        SQL = SQL & ", " & vRs!lotetraza
        TotalCajas = TotalCajas + vRs.Fields(1)
        vRs.MoveNext
    Wend
    If SQL = "" Then Exit Sub
    SQL = Mid(SQL, 2)
    'Ya tengo todos los lotes. Vere lo que hemos producido
    cad = "Select prodlin.codartic,nomartic,lotetraza,feccaduca,fhinicio from prodtrazlin,prodlin,sartic"
    cad = cad & " Where prodtrazlin.Codigo = prodlin.Codigo And prodtrazlin.idlin = prodlin.idlin And prodlin.codArtic = Sartic.codArtic"
    cad = cad & " and lotetraza in (" & SQL & ") order BY codartic,lotetraza"
    Set RN = New ADODB.Recordset
    RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not RN.EOF
        'Marzo.
        'Posicion 5 . Fecha caducidad
        cad = RN!codartic & "|" & RN!NomArtic & "|"
            
        'Subhijo
        vRs.Find "lotetraza = " & RN!lotetraza, , adSearchForward, 1
        F = "12/04/1972"
        If vRs.EOF Then
        
            MsgBox "Error lote traza: " & RN!lotetraza, vbExclamation
            cad = cad & "0|0|"
        Else
            cad = cad & vRs!lotetraza & "|" & vRs.Fields(1) & "|"
            
            If Not IsNull(RN!feccaduca) Then
                F = RN!feccaduca
            Else
                If Not IsNull(RN!fhinicio) Then
                    F = RN!fhinicio
                    F = DateAdd("m", 18, F)
                End If
            End If
        
        End If
        cad = cad & Format(F, "dd/mm/yyyy") & "|"
        Col.Add cad
        
        RN.MoveNext
    Wend
    RN.Close
    
    vRs.Close
    
    
    If PreparaImpresion Then
        'Borramos en tmp e insertamos en tmpalgo
        'Solo borramos tmp cuando es MONO impresion
        If Not MultiplesPaletsImpresion Then Conn.Execute "DELETE FROM tmppartidas  WHERE codusu = " & vUsu.Codigo
        'tmppartidas(codartic,numlote,idpartida,cantidad,idReferencia,fecha)
        SQL = ""
        For J = 1 To Col.Count
            cad = Col.Item(J)
            SQL = SQL & ", (" & vUsu.Codigo & "," & DBSet(RecuperaValor(cad, 1), "T") & "," & DBSet(RecuperaValor(cad, 2), "T")
            SQL = SQL & "," & RecuperaValor(cad, 3) & "," & DBSet(RecuperaValor(cad, 4), "N") & "," & mvarID
            SQL = SQL & "," & DBSet(RecuperaValor(cad, 5), "F") & ")"
        Next
        If SQL <> "" Then
            SQL = Mid(SQL, 2)
            cad = "INSERT INTO tmppartidas(codusu,codartic,numlote,idpartida,cantidad,idReferencia,fecha) VALUES " & SQL
            Conn.Execute cad
        End If
    End If
    
    
    
ECargaDatosPalet:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        Set Col = Nothing
    End If
    Set vRs = Nothing
    Set RN = Nothing
    
End Sub



Public Sub AñadirIdTraza(idTra As Long)
    mvaridTrazabilidad = mvaridTrazabilidad & CStr(idTra) & "|"
End Sub


'Ya tienen que haber sido creado el otro Palet
Public Sub PasarUltimasCajasA_OtroPalet(UtlimaCajaAnterior As Integer, idOtroPalet As Long, idPaletAnterior As Long, Ktrazabilidad As Long)

   'Ktrazabilidad: si en el palet hubieran mas de un lotetraza

    SQL = "UPDATE prodcajas SET idpalet= " & idOtroPalet & " WHERE"
    SQL = SQL & " idpalet =" & idPaletAnterior
    SQL = SQL & " AND lotetraza =" & Ktrazabilidad
    SQL = SQL & " AND idcaja >" & UtlimaCajaAnterior
    Conn.Execute SQL
    
    'Si hay nmas de un lote en el palet
  '  SQL = "UPDATE prodcajas SET idpalet= " & idOtroPalet & " WHERE"
  '  SQL = SQL & " idpalet =" & idPaletAnterior
  '  SQL = SQL & " AND lotetraza >" & Ktrazabilidad
  '  Conn.Execute SQL

End Sub

