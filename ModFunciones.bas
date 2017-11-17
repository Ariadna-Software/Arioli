Attribute VB_Name = "ModFunciones"
Option Explicit

Public Const ValorNulo = "Null"


Public NombreCheck As String

Public Function CompForm(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is MSComm Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 2 And Control.Name = "Text3") Or (Opcion = 3 And Control.Name = "txtAux") Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function



Public Sub limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
    
    
    
End Sub

'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim d As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else
                If InStr(1, Valor, ".") Then
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                End If
            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "H"
            Dev = "'" & Format(Valor, FormatoFecha & " hh:mm:ss") & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then
            Dev = ValorNulo
        Else
            'Modifica Laura: 04/10/05
            If vtag.TipoDato = "N" Then
                Dev = "0"
            Else
                Dev = "''"
            End If
        End If
    End If
    ValorParaSQL = Dev
End Function



Public Function InsertarDesdeForm(ByRef formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion = 0 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.Columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function



Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la funci�n.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = Cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer


    On Error GoTo EPonerCamposForma


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If TypeOf Control Is CommonDialog Then
        
        ElseIf (TypeOf Control Is TextBox) And (Control.visible = True) And (Control.Name = "Text1") Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    
                    If mTag.Columna <> "" Then
                        'Debug.Print mTag.columna
                        'If mTag.Columna = "porciva1re" Then
                        
                        campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            If mTag.TipoDato = "N" Then
                                If Val(Valor) = 0 Then
                                    Control.Text = ""
                                Else
                                    Control.Text = Valor
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.Columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.Columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    Cad = Err.Description
    Cad = "Poner campos formulario. " & vbCrLf & campo & vbCrLf & Cad & vbCrLf
    MsgBox Cad, vbExclamation
End Function



Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer

    Set mTag = New CTag
    PonerCamposFormaFrame = False


        For Each Control In formulario.Controls
        If TypeOf Control Is TextBox And Control.visible = True And Control.Name = NomTxtBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True And Control.Name = NomCheck Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.Columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox And Control.visible = True And Control.Name = NomCombo Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.Columna
                    Valor = vData.Recordset.Fields(campo)
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If

    Next Control

    'Veremos que tal
    PonerCamposFormaFrame = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function


Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.EOF) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function

'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener b�squeda. "
'End Function

'A�ado Optional CHECK As String. Para poder realizar las busquedas con los checks
Public Function ObtenerBusqueda(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim Cad As String
Dim SQL As String
Dim Tabla As String, Columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            Cad = " MAX(" & mTag.Columna & ")"
                        Else
                            Cad = " MAX({" & mTag.Tabla & "." & mTag.Columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            Cad = " MIN(" & mTag.Columna & ")"
                        Else
                            Cad = " MIN({" & mTag.Tabla & "." & mTag.Columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                    Else
                        SQL = "Select " & Cad & " from {" & mTag.Tabla & "}"
                    End If
                    SQL = ObtenerMaximoMinimo(SQL)
                    
                    Select Case mTag.TipoDato
                    Case "N"
                        If Not paraRPT Then
                            SQL = mTag.Tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
                        Else
                            SQL = "{" & mTag.Tabla & "." & mTag.Columna & "} = " & TransformaComasPuntos(SQL)
                        End If
                    Case "F"
                        If Not paraRPT Then
                            SQL = mTag.Tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Else
                            SQL = "{" & mTag.Tabla & "." & mTag.Columna & "} = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        If Not paraRPT Then
                            SQL = mTag.Tabla & "." & mTag.Columna & " = '" & SQL & "'"
                        Else
                            SQL = "{" & mTag.Tabla & "." & mTag.Columna & "} = '" & SQL & "'"
                        End If
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        SQL = mTag.Tabla & "." & mTag.Columna & " is NULL"
                    Else
                        SQL = "{" & mTag.Tabla & "." & mTag.Columna & "} is NULL"
                    End If
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            If Not paraRPT Then
                                Tabla = mTag.Tabla & "."
                            Else
                                Tabla = "{" & mTag.Tabla & "."
                            End If
                        Else
                            Tabla = ""
                        End If
                        If Not paraRPT Then
                            Columna = mTag.Columna
                        Else
                            Columna = mTag.Columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & Columna, Aux, Cad, paraRPT)
                    If Rc = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        If Not paraRPT Then
                            SQL = SQL & "(" & Cad & ")"
                        Else
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        Cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.Columna & " = " & Cad
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.Columna & "} = " & Cad
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    Else
                        Cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.Columna & " = '" & Cad & "'"
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.Columna & "} = '" & Cad & "'"
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If Control.Value = 1 Then
                        Aux = "1"
                    Else
                        If CHECK <> "" Then
                            CheckBusqueda Control
                            Tabla = NombreCheck & "|"
                            If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                        End If
                    End If
                    If Aux <> "" Then
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.Columna
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.Columna & "} "
                        End If
                        
                        Cad = Cad & " = " & Aux
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener b�squeda. "
End Function


Public Function ModificaDesdeFormulario(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUpdate As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                 cadWhere = cadWhere & "(" & mTag.Columna & " = " & Aux & ")"
    
                            Else
                                If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                                cadUpdate = cadUpdate & "" & mTag.Columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUpdate = cadUpdate & "" & mTag.Columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUpdate = cadUpdate & "" & mTag.Columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                        cadUpdate = cadUpdate & "" & mTag.Columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE Pa�sDestinatario = 'M�xico';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUpdate & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
'devuelve el valor del control vText.text formateado: 12 -> "0012"
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub

Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String
On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    I = 0
    cont = 1
    Cad = ""
    Do
        J = I + 1
        I = InStr(J, Cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                Cad = Mid(Cadena, J, I - J)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = Cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer

On Error GoTo EPonerOpcionesMenuGeneral

'A�adir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I

    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next

    On Error Resume Next

    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False

    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False

    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    
    J = Val(.mnLineas.HelpContextID)
    If J < vUsu.Nivel Then .mnLineas.Enabled = False
    
    On Error GoTo 0
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



Public Function BLOQUEADesdeFormulario(ByRef formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion <> 1 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.Columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTabla As String, cadWhere As String) As Boolean
Dim Aux As String
On Error GoTo EBloqueaRegistro

        BloqueaRegistro = False
        
        Aux = "SELECT * FROM " & cadTabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BloqueaRegistro = True
        
EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Name = "Text1" Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "Insert into zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.Tabla
        Aux = Aux & "',""" & ComprobarComillas(AuxDef) & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '�Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Private Function ComprobarComillas(Cad As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cad, """")
        If I > 0 Then
            Aux = Mid(Cad, 1, I - 1) & "\"
            Cad = Aux & Mid(Cad, I)
            J = I + 2
        End If
    Loop Until I = 0
    ComprobarComillas = Cad
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        SQL = "DELETE from zbloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.Tabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function



Public Function BloqueoManual(cadTabla As String, cadWhere As String) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWhere & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '�Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(Cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function

'Redondeo a 4 digitos
Public Function CalcularImporte4(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(Cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    vImp = Round(vImp, 4)
    CalcularImporte4 = CStr(vImp)
End Function




Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub ComprobarCobrosCliente(CodClien As String, FechaDoc As String)
'Comprueba en la tabla de Cobros Pendientes (scobro) de la Base de datos de Contabilidad
'si el cliente tiene alguna factura pendiente de cobro que ha vendido
'con fecha de vencimiento anterior a la fecha del documento: Oferta, Pedido, ALbaran,...
Dim SQL As String, vWhere As String
Dim Codmacta As String
Dim RS As ADODB.Recordset
Dim cadMen As String

    'Obtener la cuenta del cliente de la tabla sclien en Ariges
    SQL = "nomclien"
    Codmacta = DevuelveDesdeBDNew(conAri, "sclien", "codmacta", "codclien", CodClien, "N", SQL)
    If Codmacta = "" Then Exit Sub
    CodClien = CodClien & " - " & SQL
    
    'Obtener a partir de la cuenta del cliente si hay cobros pendientes en Contabilidad
    
    If vParamAplic.ContabilidadNueva Then
    
        SQL = "SELECT sum(impvenci - if(isnull(impcobro),0,impcobro))  FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        vWhere = " WHERE cobros.codmacta = '" & Codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        vWhere = vWhere & " AND recedocu=0 ORDER BY fecfactu, numfactu"
        SQL = SQL & vWhere
    
    
    Else

    
    
        SQL = "SELECT sum(impvenci - if(isnull(impcobro),0,impcobro)) FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        vWhere = " WHERE scobro.codmacta = '" & Codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
        SQL = SQL & vWhere
    
    End If
    Set RS = New ADODB.Recordset
    'Lee de la Base de Datos de CONTABILIDAD
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        If DBLet(RS.Fields(0).Value, "N") > 0 Then
            cadMen = "El Cliente tiene facturas vencidas con valor de: " & RS.Fields(0).Value & " �."
            cadMen = cadMen & vbCrLf & "�Desea Ver Detalle?"
            If MsgBox(cadMen, vbInformation + vbYesNo, "Cobros Pendientes") = vbYes Then
                'Mostrar los detalles de los cobros pendientes
                frmMensajes.cadWhere = vWhere
                frmMensajes.vCampos = CodClien
                frmMensajes.OpcionMensaje = 1
                frmMensajes.Show vbModal
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
End Sub


Public Function EsArticuloVarios(codartic As String) As Boolean
Dim Devuelve As String

    EsArticuloVarios = False
    Devuelve = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", codartic, "T")
    
    If Devuelve = "1" Or Devuelve = "2" Then 'Es Articulo de Varios y podemos modificar la Denominaci�n del Articulo
        EsArticuloVarios = True
    Else
        EsArticuloVarios = False
    End If
End Function


Public Function EsArticuloTrazabilidad(codartic As String) As Boolean
Dim Devuelve As String

    EsArticuloTrazabilidad = False
    Devuelve = DevuelveDesdeBD(conAri, "trazabilidad", "sartic", "codartic", codartic, "T")
    
    If Devuelve = "1" Then
        EsArticuloTrazabilidad = True
    Else
        EsArticuloTrazabilidad = False
    End If
End Function

Public Function EsClienteVarios(vCodClien As String) As Boolean
'Devuelve true si es un cliente de varios
Dim Devuelve As String

    EsClienteVarios = False
    Devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", vCodClien, "N")
    If Devuelve <> "" Then EsClienteVarios = CBool(Devuelve)
    'Es cliente de varios Y podemos recuperar de sclvar los datos
    'del cliente por el NIF
End Function


Public Function EsProveedorVarios(codProve As String) As Boolean
Dim Devuelve As String

    EsProveedorVarios = False
    Devuelve = DevuelveDesdeBD(conAri, "provario", "sprove", "codprove", codProve, "N")
    If Devuelve <> "" Then EsProveedorVarios = CBool(Devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function


Public Function ObtenerNSerieSiguiente(cadNSerie As String) As String
'IN -> cadNSerie: cadena con el N� Serie de Tipo: "0000-12-0011"
'OUT -> RETURN: cadena con el sig. N�Serie : "0000-12-0012"
Dim NumAux As String, numAnt As String
Dim NumAux2 As String
Dim I As Integer

    On Error Resume Next
    
    NumAux = cadNSerie
    numAnt = ""
    'Quitar los cararacter '-' y quedarse con la parte dcha
    I = InStr(1, NumAux, "-")
    While Not I = 0
        numAnt = numAnt & Mid(NumAux, 1, I)
        NumAux = Mid(NumAux, I + 1, Len(NumAux) - I)
        I = InStr(1, NumAux, "-")
    Wend
    
    If NumAux <> "" Then 'Hay q coger la parte derecha del - : 0011
        I = Len(NumAux)
        If IsNumeric(NumAux) Then
            NumAux = CStr(NumAux + 1)
            While Len(NumAux) < I
                NumAux = "0" & NumAux
            Wend
        Else
        'Coger el n� mas a la derecha, incrementarlo y concatenarlo con el principio
            NumAux2 = Mid(NumAux, I, Len(NumAux))
            While IsNumeric(NumAux2)
                I = I - 1
                NumAux2 = Mid(NumAux, I, Len(NumAux))
            Wend
            NumAux2 = Right(NumAux2, Len(NumAux2) - 1)
            numAnt = numAnt & Mid(NumAux, 1, I)
            NumAux = CStr(NumAux2 + 1)
            While Len(NumAux) < Len(NumAux2)
                NumAux = "0" & NumAux
            Wend
        End If
        
        If numAnt <> "" Then
            ObtenerNSerieSiguiente = numAnt & NumAux
        Else
            ObtenerNSerieSiguiente = NumAux
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PonerTrabajadorConectado(NomTraba As String) As String
'Pone en el campo del Form "Realizada Por" el trabajador que esta conectado en ese momento
'OUT: codTraba, NomTraba
Dim Devuelve As String

    On Error Resume Next

    NomTraba = "nomtraba"
    Devuelve = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "login", vUsu.Login, "T", NomTraba)
    If Devuelve <> "" Then
        PonerTrabajadorConectado = Format(Devuelve, "0000") 'Cod. Trabajador
    Else
        PonerTrabajadorConectado = ""
        NomTraba = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim Devuelve As String
    
    On Error Resume Next

    If codAlm = "" Then
        MsgBox "Debe introducir el Almacen.", vbInformation
    Else
        Devuelve = DevuelveDesdeBDNew(conAri, "salmpr", "codalmac", "codalmac", codAlm, "N")
        If Devuelve = "" Then
            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
            PonerAlmacen = ""
        Else
            PonerAlmacen = Format(codAlm, "000")
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'=============================================================================
'==================== REPARACIONES ===========================================

Public Sub ComprobarReparaciones(Modo As Byte, numSerie As String, codartic As String)
Dim numRep As Integer

    'Comprobar si ya esta en Reparacion
    If Modo = 3 Then ComprobarSiReparandose numSerie, codartic
    'Comprobar cuantas veces se ha reparado ya el articulo(ver historico Reparaciones)
    numRep = ComprobarNumRepHco(numSerie, codartic)
    If numRep > 0 Then
        MsgBox "Este aparato ya ha sido reparado " & numRep & " veces.", vbInformation
    End If
End Sub



Public Function ComprobarSiReparandose(numSerie As String, codartic As String) As Boolean
'Comprueba si ya el Articulo se esta reparando, es decir si existe un registro
' en la tabla scarep
'IN -> numSerie, codArtic
Dim Devuelve As String

    Devuelve = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", "numserie", numSerie, "T", , "codartic", codartic, "T")
    If Devuelve <> "" Then
        MsgBox "Este aparato ya esta en Reparaci�n.", vbInformation
        ComprobarSiReparandose = True
    Else
        ComprobarSiReparandose = False
    End If
End Function


Public Function ComprobarNumRepHco(numSerie As String, codartic As String) As Integer
'Comprueba cuantas veces se ha reparado ya el articulo
'Ver cuantos registros existen en la tabla de historico Reparaciones (schrep)
'IN -> numserie, codartic
'RETURN -> N� Reparaciones
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ENumRep

    SQL = " SELECT count(numrepar) FROM schrep "
    SQL = SQL & " WHERE numserie=" & DBSet(numSerie, "T") & " and "
    SQL = SQL & " codartic=" & DBSet(codartic, "T")

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ComprobarNumRepHco = RS.Fields(0).Value
    Else
        ComprobarNumRepHco = 0
    End If
    
    RS.Close
    Set RS = Nothing
    
ENumRep:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Function ObtenerLetraSerie(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String

    On Error Resume Next
    
    LEtra = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie = LEtra
End Function


Public Function ObtenerPoblacion(CPostal As String, ByRef provin As String) As String
'IN: "cpostal"
'OUT: en "provin" devolvemos la provincia
'     en ObtenerPoblacion devolvemos la poblacion
Dim Devuelve As String

    On Error GoTo EPoblacion

    If CPostal <> "" Then
        Devuelve = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", CPostal, "T")
        ObtenerPoblacion = Devuelve 'Nombre Poblacion
        If Devuelve <> "" Then 'Nombre Provincia
            provin = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", Mid(CPostal, 1, 2), "T")
        Else
            provin = ""
            MsgBox "No existe el CPostal " & CPostal, vbInformation
        End If
    Else
        ObtenerPoblacion = ""
        provin = ""
    End If
    
EPoblacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Poblaci�n", Err.Description
End Function


Public Sub ObtenerCtasBancoPropio(banPr As String, ctaBan As String, ctaCble As String)
'obtener la cuenta bancaria y la cuenta contable del banco propio
'(IN) banPr: cod. banco propio
'(OUT) ctaBan: cuenta bancaria
'(OUT) ctaCble: cuenta contable
Dim RS As ADODB.Recordset
Dim SQL As String

    ctaBan = ""
    ctaCble = ""

    SQL = "SELECT codbanco,codsucur,digcontr,cuentaba,codmacta"
    SQL = SQL & " from sbanpr where codbanpr=" & banPr

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        ctaBan = DBLet(RS!codbanco, "T") & "-" & DBLet(RS!codsucur, "T") & "-"
        ctaBan = ctaBan & DBLet(RS!digcontr, "T") & "-" & DBLet(RS!cuentaba, "T")
        ctaCble = DBLet(RS!Codmacta, "T")
        'obtener el nombre de la cuenta contable
        SQL = ""
        SQL = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", ctaCble, "T")
        If SQL <> "" Then ctaCble = ctaCble & "-" & SQL
    End If
    Set RS = Nothing
End Sub



Public Function ObtenerSQLcomponentes(cadWhere As String) As String
'Obtiene la consulta SQL que selecciona los articulos con n� de serie
'agrupados por tipo de articulo
Dim SQL As String

    SQL = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    SQL = SQL & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    SQL = SQL & cadWhere
    SQL = SQL & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = SQL
End Function



Public Function ComprobarStock(codartic As String, codAlmac As String, cant As String, CodTipMov As String) As Boolean
'Comprueba si el Articulo existe en el Almacen Origen y si hay
'stock suficiente para poder realizar el traspaso
Dim vStock As String
Dim vArtic As CArticulo
Dim b As Boolean

    Set vArtic = New CArticulo
    b = vArtic.Existe(codartic)
    If b Then
        b = vArtic.ExisteEnAlmacen(codAlmac, vStock)
        If b Then
            b = ComprobarHayStock(CSng(vStock), CSng(cant), codartic, vArtic.Nombre, CodTipMov)
'            If Not ComprobarHayStock(CSng(vStock), CSng(cant), codArtic, vArtic.Nombre, CodTipMov) Then
'                b = False
'            Else
'                b = True
'            End If
        End If
    End If
    Set vArtic = Nothing
    ComprobarStock = b
End Function



Public Function ObtenerPrecioSinIVAvarios(codartic As String, Precio As String) As Currency
Dim vArtic As CArticulo
Dim PreuSinIVA  As Currency

'    On Error GoTo ErrTotal
'
''    If sPorce <> "" Then curPorce = ImporteFormateado(sPorce)
'    If Precio <> "" Then PreuConIVA = ImporteFormateado(Precio) 'precio con iva

    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codartic) Then
        'precio con iva del articulo
        PreuSinIVA = vArtic.ObtenerPrecioSinIVA(Precio)
    Else
        PreuSinIVA = CCur(ComprobarCero(Precio))
    End If

'
'
'    curPorce = curPorce / 100
'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
'    Exit Function


'    Set vArtic = New CArticulo
'    If vArtic.LeerDatos(codArtic) Then
'        'precio con iva del articulo
'        PreuIVA = vArtic.ObtenerPrecioConIVA
'    End If
'
'
'    'El precio con IVA calculado a partir del importe del articulo no coincide con el
'    'precio con IVA introducido en la linea.
'    'recalculamos el importe del articulo SIN iva (se modifica precio original del artic)
'    If Round(PreuIVA, 2) <> Round(CCur(Precio), 2) Then
'        If PreuIVA <> 0 Then
'            PreuIVA = Round((vArtic.PrecioVenta * CCur(Precio)) / PreuIVA, 4)
'        Else
'            PreuIVA = Round((CCur(Precio) * 100) / (100 + vArtic.ObtenerPorceIVA), 4)
'        End If
'    Else
'        PreuIVA = vArtic.PrecioVenta
'    End If
    Set vArtic = Nothing
    ObtenerPrecioSinIVAvarios = PreuSinIVA
End Function




 



Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim Cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function CApos(Texto As String) As String
'-- (RAFA/ALZIRA) 07092006
'-- Esta funci�n procesa caracteres extra�os y de control para sentencias SQL

    Dim I As Integer
    Dim i2 As Integer
    i2 = 1
    I = InStr(i2, Texto, "'")
    While I <> 0
        Texto = Mid(Texto, 1, I) & "'" & Mid(Texto, I + 1, Len(Texto) - I)
        i2 = I + 2
        I = InStr(i2, Texto, "'")
    Wend
    i2 = 1
    I = InStr(i2, Texto, "\")
    While I <> 0
        Texto = Mid(Texto, 1, I) & "\" & Mid(Texto, I + 1, Len(Texto) - I)
        i2 = I + 2
        I = InStr(i2, Texto, "\")
    Wend

End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un n�mero."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, Cad)
  
End Function



Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function




Public Function ArticuloTieneMargen(codart As String) As Boolean
Dim Cad As String

    'Comprobar que el art�culo tiene margen comercial
    Cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", codart, "T")
    If Cad = "" Then
        Cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
        Cad = Cad & "El art�culo no tiene margen comercial para calcular nuevos precios."
        MsgBox Cad, vbExclamation
        ArticuloTieneMargen = False
        Exit Function
    End If
    
    
'    'comprobar que las tarifas del articulo tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(codArt, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El art�culo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
    
    ArticuloTieneMargen = True
    
End Function






Public Function TotalRegistros(vSQL As String, Optional vBD As Byte) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    If vBD = conConta Then 'Accede a BD de contabilidad
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub



Public Sub CheckCadenaBusqueda(ByRef CH As CheckBox, ByRef CadenaCHECKs As String)
        CheckBusqueda CH
        If InStr(1, CadenaCHECKs, NombreCheck) = 0 Then CadenaCHECKs = CadenaCHECKs & NombreCheck & "|"
End Sub




'---------------------------------------------------------------------------------
'
'       Las tabla reparaciones esta relacionada, sin FOREING KEY con
'       SAT, tipoave,trabajorealizado
'       Para saber si se puede eliminar alguno de estos
'       mantenimientos entonces trendrmos esta funcion
'
'       Opcion
'           1:  sat
'           2:  tipoave
'           3:  trabajaorealizado
Public Function SePuedeEliminarRelReparacione(Opcion As Byte, Codigo As String) As Boolean
Dim Ca As String
Dim C2 As String

    SePuedeEliminarRelReparacione = False
    If Opcion = 1 Then
        'SAT
        Ca = "codman"
    Else
        If Opcion = 2 Then
            Ca = "codavi" 'Deberia haber sido AVE de averia, no avi
        Else
            Ca = "codtrabajo"
        End If
    End If
    'Miramos primero en scarep
    C2 = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", Ca, Codigo, "N")
    If C2 <> "" Then Exit Function
        
        
    'Ahora miraremos en hco reparaciones
    C2 = DevuelveDesdeBDNew(conAri, "schrep", "numrepar", Ca, Codigo, "N")
    If C2 <> "" Then Exit Function

    
    SePuedeEliminarRelReparacione = True
End Function

Public Function SugerirCodAutomatico(marca As String, Categoria As String, modelo As String, Formato As String) As String
    '-- SugerirCodAtomatico:
    '   Esta funci�n se utiliza en el marco del par�metro descriptores y sirve, al igual que se montaba un descriptor
    '   autom�tico a partir de las descripciones de los campos de marca, categoria, modelo y formato; hacer lo propio
    '   pero con el c�digo. Con el siguiente formato
    '   MMMMCCCCmmffXXXX -> M=marca, C=categoria, m=modelo, f=formato, x=un ordinal para el c�digo
    Dim inferior As String
    Dim superior As String
    Dim comun As String
    Dim Codigo As String
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Valor As Integer
    '-- Primero trimeamos los valores por si acaso.
    marca = Left(Trim(marca) & "0000", 4)
    Categoria = Left(Trim(Categoria) & "0000", 4)
    modelo = Left(Trim(modelo) & "00", 2)
    Formato = Left(Trim(Formato) & "00", 2)
    '--
    comun = marca & Categoria & modelo & Formato
'    inferior = comun & "0000"
'    superior = comun & "9999"
'
'    SQL = "select max(codartic) from sartic where" & _
'            " codartic >= '" & inferior & "'" & _
'            " and codartic <= '" & superior & "'"
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly
'    '-- por defecto el c�digo es:
'    codigo = comun & "0001"
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            If Not IsNumeric(Right(RS.Fields(0), 4)) Then
'                MsgBox "La cola de c�digo: " & RS.Fields(0) & " no es num�rica. No puedo sugerir el c�digo siguiente", vbExclamation
'                codigo = ""
'            Else
'                Valor = Val(Right(RS.Fields(0), 4)) + 1
'                codigo = comun & Format(Valor, "0000")
'            End If
'        End If
'    End If
'    SugerirCodAutomatico = codigo
    SugerirCodAutomatico = comun
End Function

Public Function CambiaTagDescriptores(ByRef Txt As TextBox, descriptor As String) As String
    '-- Cambia el comienzo del tag del descriptor en el tag, para que cuando diga xxx no exista, aparezca
    '   la etiqueta correcta.
    Dim pos As Integer
    Dim ntag As String
    ntag = Txt.Tag
    pos = InStr(1, ntag, "|")
    If pos Then
        ntag = descriptor & Mid(ntag, pos, (Len(ntag) - pos) + 1)
    End If
    Txt.Tag = ntag
    CambiaTagDescriptores = ntag
End Function


'                                                                       CINCO DECIMALES
'Cambio MARZO 2010          4 decimales
Public Function ArticuloConTasaReciclado2(ArticuloLinea As String, ByRef Importe As Currency) As Boolean
'Public Function ArticuloConTasaReciclado2(ArticuloLinea As String, ByRef ImporteSng As Single) As Boolean
Dim RT As ADODB.Recordset
Dim SQL As String
        On Error GoTo EArticuloConTasaReciclado
        ArticuloConTasaReciclado2 = False
        SQL = "select tasareciclado from sunida,sartic where sunida.codunida =sartic.codunida and sartic.codartic=" & DBSet(ArticuloLinea, "T")
        Set RT = New ADODB.Recordset
        RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RT.EOF Then
            If Not IsNull(RT!tasareciclado) Then
                'ImporteSng = RT!tasareciclado
                Importe = RT!tasareciclado
                
                ArticuloConTasaReciclado2 = True
            End If
        End If
        RT.Close
        Set RT = Nothing
        Exit Function
EArticuloConTasaReciclado:
    MuestraError Err.Number, Err.Description, "Calculando tasa reciclado."
    Set RT = Nothing
End Function




'TarifaOferta
'--------------------------------------------------------------
'
'   A partir de PIVU (precio inicial venta unitario) aplicaremos
'   coste 1,coste 2 ,coste 3   como porcentajes
'   coste 4 coste 5  como Incremento UNITARIO
'   y finalmente le aplicaremos un margen CalculaImporteLineaTO
Public Function CalculaImporteLineaTO(PIVU As Currency, C1 As Currency, C2 As Currency, C3 As Currency, C4 As Currency, C5 As Currency, margen As Currency, LitrosUnidad As Currency) As Currency
Dim vM As Currency
Dim CUd1 As Currency
Dim CUd2 As Currency
Dim Im As Currency


            'MARZO 2010
            'FORMATEAMOS LOS PRECIOS a 3 decimales. El resultado final es el que vamos a formatear

            'Coste 1
            ' Pro ud si es menores por litro si son formatos mayores
            If LitrosUnidad > 1 Then
                
                CUd1 = Round(C4 * LitrosUnidad, 4)   '//Mayo 2009
                CUd2 = Round(C5 * LitrosUnidad, 4)   '   cambiamos el / que habia por el *
                                                     'El coste es tanto por litro para mayores si no es por UD
            Else
                CUd1 = C4
                CUd2 = C5
            End If
                
            'TEngo que reahacer todos los calculos para las obteneciones de los precios finales
            'Precio enta final UD
            Im = Round2((((PIVU * C1) + (PIVU * C2) + (PIVU * C3)) / 100) + CUd1 + CUd2, 4) + PIVU
            'Margen)
            vM = Round2((Im * margen) / 100, 4)
            
            'ANTES
            'CalculaImporteLineaTO = Im + vM
            
            'Sumamos el margen
            Im = Im + vM
            Im = Round2(Im, 3) 'Redondeamos a 3 decimales MARZO 2010
            CalculaImporteLineaTO = Im
            
            
End Function

'Julio 2014
'Si el numero de lote es igual al ofertado, entonces lo incremento
Public Sub IncrementaLoteCompras(NumeroLote As String)
Dim NumeroL As Long
Dim Mc As CTiposMov



    On Error GoTo eIncrementaLoteCompras

        Set Mc = New CTiposMov
        If UCase(Right(NumeroLote, 2)) = "-A" Then
            Mc.Leer "LOA"
            NumeroL = Val(Mid(NumeroLote, 1, Len(NumeroLote) - 2))
        Else
            Mc.Leer "LOT"
            NumeroL = Val(NumeroLote)
        End If
        If NumeroL = Mc.contador Then Mc.IncrementarContador Mc.TipoMovimiento
        Set Mc = Nothing

        
    Exit Sub
eIncrementaLoteCompras:
    MuestraError Err.Number, "Guardano ultimo lote ", Err.Description
    Set Mc = Nothing
End Sub
