VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Servidor donde esta MYSQL
Private mvarBD As String 'Variable local

'---- Usuario general BD
Private mvarCOMM As String 'Variable local

'---- Password usuario general BD
Private mvarVelocidad As String 'Variable local

'---- Segundos protector
Private mvarSegundos As String 'Variable local



'Variables de este modulo
Dim NF As Integer
Dim Mens As String
'variables locales para almacenar los valores de las propiedades
Private mvarAplicacion As String 'copia local

Public Property Let BD(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Aplicacion = 5
    mvarBD = vData
End Property


Public Property Get BD() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Aplicacion
    BD = mvarBD
End Property


Public Property Let kCOMM(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarCOMM = vData
End Property

Public Property Get kCOMM() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     kCOMM = mvarCOMM
End Property

'----
'---- Usuario general BD
'----
Public Property Let Velocidad(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarVelocidad = vData
End Property

Public Property Get Velocidad() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     Velocidad = mvarVelocidad
End Property




Public Property Let Segundos(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarSegundos = vData
End Property

Public Property Get Segundos() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     Segundos = mvarSegundos
End Property







'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim nFile As Integer
    
    On Error GoTo Err1
    Fichero = App.Path & "\ConfigBas.ini"
    
    nFile = FreeFile
    Open Fichero For Input As #nFile   ' Abre el archivo.
    Do While Not EOF(nFile) ' Repite el bucle hasta el final del archivo.
        Line Input #nFile, Lin
        Select Case Lin
           Case "[BD]", "[COMM]", "[VELOCIDAD]", "[SEG.PROTEC]", "[VACIO]"
               
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[BD]": mvarBD = ValParam
                    Case "[COMM]": mvarCOMM = ValParam
                    Case "[VELOCIDAD]": mvarVelocidad = ValParam
                    Case "[SEG.PROTEC]": mvarSegundos = ValParam
                    Case "[VACIO]":  ValParam = ""
                   
'                    Case "[PEDIR]": mvarPedirPasswd = CBool(ValParam)
'                    Case "[INTEG]": mvarIntegraciones = ValParam
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #nFile    ' Cierra el archivo.
    
    
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function




'------------------------------------------------------
'     GUARDAR

Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim nFile As Integer

    On Error GoTo Err1
    Fichero = App.Path & "\ConfigBas.ini"
    nFile = FreeFile
    Open Fichero For Output As #nFile   ' Abre el archivo.
    Print #nFile, "[BD]"
    Print #nFile, mvarBD
    Print #nFile, "<-------->"
    Print #nFile, "[COMM]"
    Print #nFile, mvarCOMM
    Print #nFile, "<-------->"
    Print #nFile, "[VELOCIDAD]"
    Print #nFile, mvarVelocidad
    Print #nFile, "<-------->"
    Print #nFile, "[SEG.PROTEC]"
    Print #nFile, mvarSegundos
    Print #nFile, "<-------->"

    
'    Print #NFile, "[PEDIR]"
'    Print #NFile, mvarPedirPasswd
'    Print #NFile, "<-------->"
'    Print #NFile, "[INTEG]"
'    Print #NFile, mvarIntegraciones
'    Print #NFile, "<-------->"
       
    Close #nFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function



