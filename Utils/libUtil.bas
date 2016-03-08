Attribute VB_Name = "libUtil"
Option Explicit



Public Const NumeroDeDecimales = 2


'Formato de fecha
Public FormatoFecha As String
Public FormatoFechaHora As String

    '#,##0.00
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(10,4)
Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoCantidad2 As String 'Decimal(8,2)
Public FormatoDescuento As String 'Decimal(4,2)
Public FormatoKms As String 'Decimal(8,4)
Public FormatoPorcen As String 'Decimal(5,2)

Public CadenaDesdeOtroForm As String


'Conexión a la BD Ariges de la empresa
Public Conn As ADODB.Connection





Public Sub Main()
Dim cad As String
    If Dir(App.Path & "\aqui.txt", vbArchive) = "" Then
        cad = InputBox("Base datos nº: ", "Conexion", "1")
    Else
        cad = InputBox("Base datos nº: ", "Conexion", "9")
    End If
    
    If Not IsNumeric(cad) Then cad = "1"
    If Not AbrirConexion(cad) Then End
    
    OtrasAcciones
    
    
    frmPpUtuil.Show
End Sub



Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "###,##0.0000"  'Decimal(10,4)
    
    'Por si acasomcambaimos la aplicacion los numeros de decimales
    'ANTES
    'FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    'FormatoCantidad2 = "###,##0.00"   'Decimal(8,2)
    'Ahora
    FormatoCantidad = "##,###,##0." & String(NumeroDeDecimales, "0")
    FormatoCantidad2 = "###,##0." & String(NumeroDeDecimales, "0")
    
    FormatoDescuento = "#0.00" 'Decima(4,2)
    FormatoKms = "#,##0.00##" 'Decimal(8,4)
    FormatoPorcen = "##0.00" 'Decima(5,2)
    

    
End Sub


'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion(Kariges As String) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection

    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    

    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;DATABASE=Ariges" & Kariges
    cad = cad & ";UID=root"
    cad = cad & ";PWD=aritel"
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Ariges.", Err.Description
End Function





Public Sub MuestraError(numero As Long, Optional Cadena As String, Optional Desc As String)
    Dim cad As String
    Dim aUX As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        cad = cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    aUX = ""

    If aUX <> "" Then Desc = aUX
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If aUX = "" Then cad = cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox cad, vbExclamation
End Sub


Public Function Espera(Segundos As Single)
Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function





'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, Cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(Cadena, J, I - J)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(Cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "," & Mid(Cadena, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = Cadena
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, Cadena, ",")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "." & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = Cadena
End Function






Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function

Public Function EsFechaOKTex(ByRef T As TextBox) As Boolean
Dim Te As String
    
    
    EsFechaOKTex = False
    Te = T.Text
    If EsFechaOK(Te) Then
        T.Text = Te
        EsFechaOKTex = True
    Else
       T.Text = ""
    End If
End Function


Public Function EsFechaOK(T As String) As Boolean
Dim cad As String
Dim mes As String, dia As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
       'debe ser una cadena tipo:020105 y la convertimos a 02/01/05
       If Not IsNumeric(cad) Then
            EsFechaOK = False
            Exit Function
       End If
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(cad, 1, 2)
        If dia < 1 Or dia > 31 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(cad, 3, 2)
        If mes < 1 Or mes > 12 Then
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    Else
        dia = Mid(cad, 1, 2)
        mes = Mid(cad, 4, 2)
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T = Format(cad, "dd/mm/yyyy")
      '==== Añade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function


Public Sub PonerFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ObtenerFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SelStart = 0
    T.SelLength = Len(T.Text)
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = "0"
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function





Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim aUX As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    

    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function


Public Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub
