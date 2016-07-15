Attribute VB_Name = "libBascula"
Option Explicit

Public Conn As Connection

Public Config As cConfigurar







Public Sub Main()

    Set Config = New cConfigurar
    If Config.Leer = 1 Then
    
        'Valore por defecto
        Config.BD = 1
        Config.kCOMM = 4
        Config.Velocidad = 9600
    
        Config.Grabar
        End
    End If
    
    If Not AbrirConexion Then Exit Sub
    
    Form1.Show
    
    
End Sub




Public Function AbrirConexion() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

   
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;DATABASE=ariges" & Config.BD
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
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        cad = cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    ''Numeros de errores que contolamos
    'If Conn.Errors.Count > 0 Then
    '    ControlamosError Aux
    '    Conn.Errors.Clear
    'Else
    '    Aux = ""
    'End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then cad = cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox cad, vbExclamation
End Sub



Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


