Attribute VB_Name = "libPoste"
Option Explicit
Public EsServicio As Boolean

Public Conn As Connection

Private SegundosServicio As Byte
Private Reintentos As Byte

Public Function AbrirConexion2(Kariges2 As Byte) As String
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion2 = "Abriendo BD"
    Set Conn = Nothing
    Set Conn = New Connection

    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;DATABASE=Ariges" & CStr(Kariges2)
   ' cad = cad & ";UID=root"
   ' cad = cad & ";PWD=aritel"
    Cad = Cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = Cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion2 = ""
    Exit Function
    
EAbrirConexion:
    AbrirConexion2 = "Abrir conexión BD:Ariges." & vbCrLf & Err.Description
End Function

Public Sub Main()

    Load frmTerminal
    If Not EsServicio Then frmTerminal.Show
End Sub


Public Function DevNombreSQL(Cadena As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = Cadena
End Function



Public Function Ejecuta(SQL As String) As Boolean
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        Ejecuta = False
    Else
        Ejecuta = True
    End If
End Function


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




'Icono: 0- Informacion  1.- Excalmacion   2.-Crititcal
Public Sub MensajeError(Cadena As String, Icono As Byte)
    If EsServicio Then
        '
        If Icono = 0 Then
            Icono = 4
        ElseIf Icono = 1 Then
            Icono = 2
        Else
            Icono = 1
        End If
        App.LogEvent Cadena, Icono
    Else
        If Icono = 0 Then
            Icono = 64
        ElseIf Icono = 1 Then
            Icono = 48
        Else
            Icono = 16
        End If
        MsgBox Cadena, Icono
    End If
    Err.Clear  'hay que borrar el error
End Sub

